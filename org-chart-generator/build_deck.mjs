import fs from "node:fs/promises";
import path from "node:path";
import { Presentation, PresentationFile } from "@oai/artifact-tool";

const SETTINGS_PATH = process.env.ORG_SETTINGS;
const DATA_PATH = process.env.ORG_DATA;
if (!SETTINGS_PATH || !DATA_PATH) {
  throw new Error("ORG_SETTINGS and ORG_DATA must be set. Run this generator through run.ps1.");
}

const settingsPath = path.resolve(SETTINGS_PATH);
const baseDir = path.dirname(settingsPath);
const settings = JSON.parse(await fs.readFile(settingsPath, "utf8"));
const data = JSON.parse(await fs.readFile(path.resolve(DATA_PATH), "utf8"));
const options = settings.options ?? {};
const slideWidth = Number(options.slideWidth ?? 1280);
const slideHeight = Number(options.slideHeight ?? 720);
const fontFamily = options.fontFamily || "Aptos";
const photoExtensions = (options.photoExtensions ?? [".jpg", ".jpeg", ".png", ".webp"])
  .map((extension) => extension.toLowerCase());

function resolveSettingPath(value) {
  if (!value) throw new Error("A required path is blank in the settings file");
  return path.isAbsolute(value) ? path.normalize(value) : path.resolve(baseDir, value);
}

const photoFolder = resolveSettingPath(settings.photoFolder);
const outputPath = resolveSettingPath(settings.outputPath);

function normalizeId(value) {
  return String(value ?? "").trim().toUpperCase();
}

function initials(name) {
  return String(name || "?")
    .split(/\s+/)
    .filter(Boolean)
    .slice(0, 2)
    .map((part) => part[0].toUpperCase())
    .join("") || "?";
}

function safeColor(value, fallback = "#5B7FA3") {
  const text = String(value ?? "").trim();
  return /^#[0-9a-f]{6}$/i.test(text) ? text.toUpperCase() : fallback;
}

function textColorFor(background) {
  const hex = safeColor(background).slice(1);
  const red = Number.parseInt(hex.slice(0, 2), 16);
  const green = Number.parseInt(hex.slice(2, 4), 16);
  const blue = Number.parseInt(hex.slice(4, 6), 16);
  return 0.299 * red + 0.587 * green + 0.114 * blue > 155 ? "#172033" : "#FFFFFF";
}

function makePhotoIndex(files) {
  const index = new Map();
  for (const file of files) {
    const extension = path.extname(file.name).toLowerCase();
    if (!file.isFile() || !photoExtensions.includes(extension)) continue;
    const stem = path.basename(file.name, extension);
    const tokens = stem.match(/[A-Za-z0-9]{4}/g) ?? [];
    for (const token of tokens) {
      const id = normalizeId(token);
      if (!index.has(id)) index.set(id, path.join(photoFolder, file.name));
    }
  }
  return index;
}

let photoDirectory;
try {
  photoDirectory = await fs.readdir(photoFolder, { withFileTypes: true });
} catch (error) {
  throw new Error(`Unable to read photo folder '${photoFolder}': ${error.message}`);
}
const photoIndex = makePhotoIndex(photoDirectory);

function prepareState(records, stateName) {
  const employeesById = new Map();
  const teams = new Map();
  for (const raw of records) {
    const employee = {
      id: normalizeId(raw.employeeId),
      name: String(raw.employeeName).trim(),
      managerId: normalizeId(raw.managerId),
      managerName: String(raw.managerName).trim(),
      role: String(raw.role || "Unspecified").trim(),
      sourceRow: raw.sourceRow,
    };
    employeesById.set(employee.id, employee);
    if (!teams.has(employee.managerId)) {
      teams.set(employee.managerId, {
        id: employee.managerId,
        name: employee.managerName,
        members: [],
        stateName,
      });
    }
    const team = teams.get(employee.managerId);
    if (team.name.toLocaleLowerCase() !== employee.managerName.toLocaleLowerCase()) {
      throw new Error(`Manager ${employee.managerId} has inconsistent names in ${stateName}`);
    }
    team.members.push(employee);
  }
  for (const team of teams.values()) {
    const managerRecord = employeesById.get(team.id);
    team.managerRole = managerRecord?.role || "Manager";
    team.managerPhotoId = team.id;
    if (options.includeManagersInEmployeeCount && managerRecord && !team.members.some((member) => member.id === team.id)) {
      team.members.unshift(managerRecord);
    }
    team.members.sort((a, b) => a.role.localeCompare(b.role) || a.name.localeCompare(b.name));
  }
  return { employeesById, teams };
}

const current = prepareState(data.current, "current");
const future = prepareState(data.future, "future");
const allManagerIds = new Set([...current.teams.keys(), ...future.teams.keys()]);

function inferDirectorId() {
  const configured = normalizeId(options.directorId);
  if (configured) {
    if (!allManagerIds.has(configured)) {
      throw new Error(`Configured directorId '${configured}' is not a manager in either sheet`);
    }
    return configured;
  }
  const managedManagerIds = new Set();
  for (const state of [current, future]) {
    for (const employee of state.employeesById.values()) {
      if (allManagerIds.has(employee.id)) managedManagerIds.add(employee.id);
    }
  }
  const roots = [...allManagerIds].filter((id) => !managedManagerIds.has(id));
  if (roots.length === 1) return roots[0];
  const directReportCounts = new Map([...allManagerIds].map((id) => [id, 0]));
  for (const state of [current, future]) {
    for (const employee of state.employeesById.values()) {
      if (allManagerIds.has(employee.id) && directReportCounts.has(employee.managerId)) {
        directReportCounts.set(employee.managerId, directReportCounts.get(employee.managerId) + 1);
      }
    }
  }
  return [...allManagerIds].sort((a, b) =>
    (directReportCounts.get(b) - directReportCounts.get(a)) || a.localeCompare(b)
  )[0];
}

const directorId = inferDirectorId();

function reportingParent(managerId) {
  const employee = future.employeesById.get(managerId) ?? current.employeesById.get(managerId);
  return employee?.managerId && allManagerIds.has(employee.managerId) ? employee.managerId : null;
}

function managerOrder() {
  const children = new Map([...allManagerIds].map((id) => [id, []]));
  for (const managerId of allManagerIds) {
    const parent = reportingParent(managerId);
    if (parent && parent !== managerId) children.get(parent).push(managerId);
  }
  for (const list of children.values()) {
    list.sort((a, b) => managerName(a).localeCompare(managerName(b)));
  }
  const ordered = [];
  const visited = new Set();
  const queue = [directorId];
  while (queue.length) {
    const id = queue.shift();
    if (!id || visited.has(id)) continue;
    visited.add(id);
    ordered.push(id);
    queue.push(...(children.get(id) ?? []));
  }
  const remaining = [...allManagerIds]
    .filter((id) => !visited.has(id))
    .sort((a, b) => managerName(a).localeCompare(managerName(b)));
  return [...ordered, ...remaining];
}

function managerName(id) {
  return future.teams.get(id)?.name ?? current.teams.get(id)?.name ?? id;
}

function addText(slide, text, position, style = {}) {
  const shape = slide.shapes.add({
    geometry: "textbox",
    position,
    fill: "none",
    line: { fill: "none", width: 0 },
  });
  shape.text = String(text);
  shape.text.style = {
    typeface: fontFamily,
    fontSize: style.fontSize ?? 18,
    bold: style.bold ?? false,
    color: style.color ?? "#172033",
    alignment: style.alignment ?? "left",
    verticalAlignment: style.verticalAlignment ?? "middle",
    autoFit: "shrinkText",
  };
  return shape;
}

function addRect(slide, position, fill, lineFill = fill, radius = "roundRect") {
  return slide.shapes.add({
    geometry: radius,
    position,
    fill,
    line: { fill: lineFill, width: 1 },
  });
}

const missingPhotos = new Set();

async function addPortrait(slide, person, position, accentColor, labelSize = 14) {
  const photoPath = photoIndex.get(person.id);
  if (photoPath) {
    const bytes = new Uint8Array(await fs.readFile(photoPath));
    slide.images.add({
      blob: bytes,
      contentType: ({ ".jpg": "image/jpeg", ".jpeg": "image/jpeg", ".png": "image/png", ".webp": "image/webp" })[path.extname(photoPath).toLowerCase()],
      alt: `${person.name} (${person.id})`,
      fit: "cover",
      geometry: "ellipse",
      position,
    });
  } else {
    missingPhotos.add(`${person.id}\t${person.name}`);
    addRect(slide, position, accentColor, accentColor, "ellipse");
    addText(slide, initials(person.name), position, {
      fontSize: Math.max(12, position.width * 0.34),
      bold: true,
      color: textColorFor(accentColor),
      alignment: "center",
    });
  }
  if (labelSize > 0) {
    const parts = [];
    if (options.showEmployeeNames !== false) parts.push(person.name);
    if (options.showEmployeeIds === true) parts.push(person.id);
    if (parts.length) {
      addText(slide, parts.join("\n"), {
        left: position.left - 6,
        top: position.top + position.height + 3,
        width: position.width + 12,
        height: labelSize * 2.8,
      }, { fontSize: labelSize, alignment: "center", color: "#303A4A" });
    }
  }
}

function roleColor(role) {
  const configured = settings.roleColors ?? {};
  const direct = configured[role];
  if (direct) return safeColor(direct);
  const key = Object.keys(configured).find((candidate) => candidate.toLowerCase() === role.toLowerCase());
  return safeColor(key ? configured[key] : configured.default);
}

function groupByRole(members) {
  const groups = new Map();
  for (const member of members) {
    if (!groups.has(member.role)) groups.set(member.role, []);
    groups.get(member.role).push(member);
  }
  return [...groups.entries()].sort(([a], [b]) => a.localeCompare(b));
}

function formatDelta(delta) {
  if (delta > 0) return `+${delta}`;
  if (delta < 0) return String(delta);
  return "0";
}

function slideTitle(team, stateLabel) {
  const stateTitle = stateLabel === "current" ? settings.titles?.current : settings.titles?.future;
  return `${team.name}, ${stateTitle || stateLabel}`;
}

async function addTeamSlide(presentation, managerId, stateLabel) {
  const state = stateLabel === "current" ? current : future;
  const team = state.teams.get(managerId) ?? {
    id: managerId,
    name: managerName(managerId),
    managerRole: "Manager",
    members: [],
  };
  const currentCount = current.teams.get(managerId)?.members.length ?? 0;
  const futureCount = future.teams.get(managerId)?.members.length ?? 0;
  const count = team.members.length;
  const maxEmployees = Number(options.maximumEmployeesPerManager ?? 60);
  if (count > maxEmployees) {
    throw new Error(`${team.name} (${managerId}) has ${count} employees in ${stateLabel}, above maximumEmployeesPerManager ${maxEmployees}`);
  }

  const slide = presentation.slides.add();
  slide.background.fill = "#F5F7FA";

  addText(slide, slideTitle(team, stateLabel), { left: 48, top: 24, width: 920, height: 52 }, {
    fontSize: 32,
    bold: true,
    color: "#172033",
  });

  const totalText = stateLabel === "future"
    ? `${count} employees  (${formatDelta(futureCount - currentCount)} vs current)`
    : `${count} employees`;
  addText(slide, totalText, { left: 960, top: 26, width: 270, height: 48 }, {
    fontSize: 20,
    bold: true,
    color: stateLabel === "future" && futureCount !== currentCount
      ? (futureCount > currentCount ? "#1B7F45" : "#B54237")
      : "#445064",
    alignment: "right",
  });

  const managerAccent = managerId === directorId ? "#203864" : "#44546A";
  addRect(slide, { left: 48, top: 94, width: 1184, height: 118 }, "#FFFFFF", "#D9E0E8");
  await addPortrait(slide, { id: managerId, name: team.name }, { left: 74, top: 108, width: 88, height: 88 }, managerAccent, 0);
  addText(slide, team.name, { left: 186, top: 112, width: 650, height: 44 }, {
    fontSize: 27,
    bold: true,
  });
  addText(slide, `${team.managerRole}${managerId === directorId ? "  Director" : ""}\nEmployee ID ${managerId}`, {
    left: 186, top: 154, width: 650, height: 44,
  }, { fontSize: 16, color: "#5C6778" });

  const roleGroups = groupByRole(team.members);
  if (!roleGroups.length) {
    addText(slide, "No employees assigned to this manager in this state", {
      left: 48, top: 300, width: 1184, height: 80,
    }, { fontSize: 24, color: "#6B7280", alignment: "center" });
    return;
  }

  const contentTop = 234;
  const contentBottom = 690;
  const availableHeight = contentBottom - contentTop;
  const gap = 12;
  const groupWidth = (1184 - gap * (roleGroups.length - 1)) / roleGroups.length;
  const minimumTileWidth = Number(options.minimumEmployeeTileWidth ?? 54);
  if (groupWidth < Math.max(160, minimumTileWidth * 1.7)) {
    throw new Error(`${team.name} has ${roleGroups.length} role groups. Increase slide width or consolidate roles to keep one slide legible.`);
  }

  for (let groupIndex = 0; groupIndex < roleGroups.length; groupIndex += 1) {
    const [role, members] = roleGroups[groupIndex];
    const color = roleColor(role);
    const left = 48 + groupIndex * (groupWidth + gap);
    addRect(slide, { left, top: contentTop, width: groupWidth, height: availableHeight }, "#FFFFFF", "#D9E0E8");
    addRect(slide, { left, top: contentTop, width: groupWidth, height: 48 }, color, color);
    addText(slide, `${role}  ${members.length}`, { left: left + 10, top: contentTop + 2, width: groupWidth - 20, height: 43 }, {
      fontSize: 18,
      bold: true,
      color: textColorFor(color),
      alignment: "center",
    });

    const innerLeft = left + 14;
    const innerTop = contentTop + 62;
    const innerWidth = groupWidth - 28;
    const innerHeight = availableHeight - 76;
    const tileGap = 10;
    let columns = Math.max(1, Math.floor((innerWidth + tileGap) / (minimumTileWidth + tileGap)));
    columns = Math.min(columns, members.length);
    let rows = Math.ceil(members.length / columns);
    let tileWidth = (innerWidth - tileGap * (columns - 1)) / columns;
    let photoSize = Math.min(tileWidth, 78);
    const labelHeight = options.showEmployeeNames === false && options.showEmployeeIds !== true ? 0 : 36;
    let tileHeight = photoSize + labelHeight + 8;
    while (rows * tileHeight + (rows - 1) * tileGap > innerHeight && columns < members.length) {
      columns += 1;
      rows = Math.ceil(members.length / columns);
      tileWidth = (innerWidth - tileGap * (columns - 1)) / columns;
      photoSize = Math.min(tileWidth, 78);
      tileHeight = photoSize + labelHeight + 8;
    }
    if (tileWidth < minimumTileWidth || rows * tileHeight + (rows - 1) * tileGap > innerHeight) {
      throw new Error(`${team.name}, role '${role}' has ${members.length} employees and cannot fit legibly on one slide. Raise slide dimensions, lower minimumEmployeeTileWidth, or consolidate the data.`);
    }
    const actualGridWidth = columns * tileWidth + (columns - 1) * tileGap;
    const gridLeft = innerLeft + (innerWidth - actualGridWidth) / 2;
    for (let index = 0; index < members.length; index += 1) {
      const row = Math.floor(index / columns);
      const column = index % columns;
      const photoLeft = gridLeft + column * (tileWidth + tileGap) + (tileWidth - photoSize) / 2;
      const photoTop = innerTop + row * (tileHeight + tileGap);
      await addPortrait(slide, members[index], {
        left: photoLeft,
        top: photoTop,
        width: photoSize,
        height: photoSize,
      }, color, Math.max(9, Math.min(12, tileWidth * 0.17)));
    }
  }
}

const presentation = Presentation.create({ slideSize: { width: slideWidth, height: slideHeight } });
const orderedManagers = managerOrder();
for (const managerId of orderedManagers) {
  await addTeamSlide(presentation, managerId, "current");
  await addTeamSlide(presentation, managerId, "future");
}

if (options.failOnMissingPhotos && missingPhotos.size) {
  throw new Error(`Missing photos for ${missingPhotos.size} people. See the report after rerunning with failOnMissingPhotos=false.`);
}

await fs.mkdir(path.dirname(outputPath), { recursive: true });
await (await PresentationFile.exportPptx(presentation)).save(outputPath);
const reportPath = outputPath.replace(/\.pptx$/i, "") + ".missing-photos.txt";
const report = missingPhotos.size
  ? `Missing photos (${missingPhotos.size})\n\n${[...missingPhotos].sort().join("\n")}\n`
  : "No missing photos.\n";
await fs.writeFile(reportPath, report, "utf8");
console.log(`Created ${outputPath}`);
console.log(`Slides: ${orderedManagers.length * 2}`);
console.log(`Missing-photo report: ${reportPath}`);
