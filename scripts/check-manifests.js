const fs = require("fs");

const DEV_MANIFEST = "manifest.dev.xml";
const PRODUCTION_MANIFEST = "manifest.xml";
const PACKAGE_JSON = "package.json";
const PACKAGE_LOCK = "package-lock.json";
const DEV_URL = "https://localhost:3001";
const PRODUCTION_URL = "https://tool2.wh-sv.de";

function readManifest(path) {
  if (!fs.existsSync(path)) {
    throw new Error(`${path} fehlt.`);
  }

  return fs.readFileSync(path, "utf8");
}

function assertContains(content, expected, path) {
  if (!content.includes(expected)) {
    throw new Error(`${path} enthaelt nicht ${expected}.`);
  }
}

function assertDoesNotContain(content, forbidden, path) {
  if (content.includes(forbidden)) {
    throw new Error(`${path} enthaelt unerwartet ${forbidden}.`);
  }
}

function readJson(path) {
  return JSON.parse(readManifest(path));
}

const localManifest = readManifest(DEV_MANIFEST);
const productionManifest = readManifest(PRODUCTION_MANIFEST);
const packageVersion = readJson(PACKAGE_JSON).version;
const packageLock = readJson(PACKAGE_LOCK);

assertContains(localManifest, DEV_URL, DEV_MANIFEST);
assertDoesNotContain(localManifest, PRODUCTION_URL, DEV_MANIFEST);

assertContains(productionManifest, PRODUCTION_URL, PRODUCTION_MANIFEST);
assertDoesNotContain(productionManifest, DEV_URL, PRODUCTION_MANIFEST);

if (localManifest === productionManifest) {
  throw new Error("DEV- und Produktionsmanifest duerfen nicht identisch sein.");
}

const devVersionMatch = localManifest.match(/<Version>([^<]+)<\/Version>/);
const prodVersionMatch = productionManifest.match(/<Version>([^<]+)<\/Version>/);

if (!devVersionMatch || !prodVersionMatch) {
  throw new Error("Manifest-Versionen konnten nicht gelesen werden.");
}

if (devVersionMatch[1] !== prodVersionMatch[1]) {
  throw new Error("DEV- und Produktionsmanifest muessen dieselbe Version haben.");
}

if (devVersionMatch[1] !== packageVersion) {
  throw new Error("package.json und Manifeste muessen dieselbe Version haben.");
}

if (packageLock.version !== packageVersion) {
  throw new Error("package-lock.json und package.json muessen dieselbe Version haben.");
}

if (!packageLock.packages || !packageLock.packages[""] || packageLock.packages[""].version !== packageVersion) {
  throw new Error("package-lock.json package version ist nicht synchron.");
}

const devIdMatch = localManifest.match(/<Id>([^<]+)<\/Id>/);
const prodIdMatch = productionManifest.match(/<Id>([^<]+)<\/Id>/);

if (!devIdMatch || !prodIdMatch) {
  throw new Error("Manifest-IDs konnten nicht gelesen werden.");
}

if (devIdMatch[1] === prodIdMatch[1]) {
  throw new Error("DEV- und Produktionsmanifest muessen unterschiedliche App-IDs haben.");
}

console.log("Manifest-Struktur ok.");
