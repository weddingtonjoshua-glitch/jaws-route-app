"use client";
import { useState, useEffect, useRef, useCallback } from "react";
import * as XLSX from "xlsx";

const sleep = ms => new Promise(r => setTimeout(r, ms));

async function geocodeAddress(address, defaultCity = "") {
  try {
    const hasCity = /,\s*\w/.test(address) || /\b(FL|Destin|Miramar|Santa Rosa|Inlet Beach|30A|Navarre|Fort Walton|Niceville|Seaside|Rosemary)\b/i.test(address);
    const fullAddress = hasCity ? address : `${address}, ${defaultCity}`;
    const q = encodeURIComponent(fullAddress);
    const res = await fetch(`https://nominatim.openstreetmap.org/search?q=${q}&format=json&limit=1&countrycodes=us`);
    const data = await res.json();
    if (data?.[0]) return { lat: +data[0].lat, lng: +data[0].lon, resolvedAs: fullAddress };
  } catch {}
  return null;
}

// Built-in coordinate estimator for 30A / Destin corridor — no API needed
function estimateCoords(address) {
  const a = address.toLowerCase();

  // ── House-number interpolation for major roads ──────────────────────────────

  // Hidden Lake Way (#97–504 runs E→W through WaterColor/WaterSound)
  const hlw = a.match(/^(\d+)\s+hidden lake way/);
  if (hlw) {
    const n = parseInt(hlw[1]);
    return { lat: 30.355, lng: -86.09 - Math.max(0, n - 97) * 0.000122, resolvedAs: address };
  }

  // Cassine Garden Cir (#31–352)
  const cg = a.match(/^(\d+)\s+cassine garden/);
  if (cg) {
    const n = parseInt(cg[1]);
    return { lat: 30.360, lng: -86.150 - Math.max(0, n - 31) * 0.0000562, resolvedAs: address };
  }

  // Gulf Cove Ct (#12–82, tight cluster)
  if (a.includes("gulf cove")) return { lat: 30.345, lng: -86.215, resolvedAs: address };

  // Scenic Hwy 98 / Scenic Gulf — house number → longitude
  const h98 = a.match(/^(\d{4,})\s+scenic\s+(hwy|highway|gulf)/);
  if (h98) {
    const n = parseInt(h98[1]);
    const lng = -86.38 - (n - 1987) * 0.0000661;
    return { lat: 30.385, lng: Math.max(-86.56, lng), resolvedAs: address };
  }

  // W County Hwy 30A — house number → longitude
  const h30a = a.match(/^(\d{4,})\s+w[\s.]+co/);
  if (h30a) {
    const n = parseInt(h30a[1]);
    return { lat: 30.370, lng: Math.max(-86.42, -86.22 - (n - 2000) * 0.0000380), resolvedAs: address };
  }

  // ── City / area detection (check most-specific first) ─────────────────────

  // Destin — check before anything else that might false-match
  const destinMatch = [
    "destiny way","bonaire cay","papaya park","kono way","pritchard",
    "ocean view dr","luke ave","tarpon","gulf shore drive",
  ].some(p => a.includes(p)) || a.includes("32541") || a.includes(", destin");
  if (destinMatch) {
    const lng = (a.includes("bonaire") || a.includes("papaya") || a.includes("destiny")) ? -86.52
      : (a.includes("kono") || a.includes("luke ave") || a.includes("tarpon")) ? -86.50
      : -86.48;
    return { lat: 30.390, lng, resolvedAs: address };
  }

  // Seacrest / Inlet Beach
  const seacrestMatch = ["pelican glide","inlet beach","cobia run","seacrest beach blvd",
    "patina blvd","sandy shores","winston lane"].some(p => a.includes(p))
    || a.includes("32461") || a.includes("seacrest");
  if (seacrestMatch) {
    return { lat: 30.310, lng: (a.includes("pelican glide") || a.includes("inlet beach")) ? -86.00 : -86.04, resolvedAs: address };
  }

  // Miramar Beach — split into eastern/central sub-areas
  const miramarMatch = a.includes("32550") || a.includes("miramar beach") || a.includes(", miramar");
  if (miramarMatch) {
    const eastern = ["st simon cir","lake ct","payne st","st francis","sandy dunes cir",
      "rue caribe","ballamore cove","monaco","snapper","cypress passage","windrift"].some(p => a.includes(p));
    return { lat: 30.385, lng: eastern ? -86.37 : -86.42, resolvedAs: address };
  }

  // Santa Rosa Beach — street-level lookup, ordered E→W
  const srbMatch = a.includes("32459") || a.includes("santa rosa");
  if (srbMatch) {
    const srbTable = [
      [["beachfront trail","hiker st","north ryan","emerald dunes"], -86.08],
      [["chelsea loop","sugar sand","sawgrass","cabana trail","shady pines"], -86.16],
      [["daybreak"], -86.20],
      [["seawinds","seaward dr","sand dunes rd","seapointe","sandalwood","high dune","dune rosemary"], -86.24],
      [["sundown","dune side lane","longue vue","gulf ridge","holly st","canal st",
        "pointe circle","gulf view","beachwalk","dothan","morgans trail","n grande beach",
        "emerald beach circle","palmeira","s co"], -86.28],
      [["blue crab","ventana"], -86.33],
      [["155 baird","baird rd"], -86.30],
    ];
    for (const [pats, lng] of srbTable) {
      if (pats.some(p => a.includes(p))) return { lat: 30.340, lng, resolvedAs: address };
    }
    return { lat: 30.340, lng: -86.22, resolvedAs: address };
  }

  return { lat: 30.380, lng: -86.30, resolvedAs: address };
}

const C = {
  bg: "#F8F9FA", surface: "#FFFFFF", s2: "#F1F3F5",
  border: "#E0E3E7", amber: "#F59E0B", amberBg: "#FFFBEB",
  amberBorder: "#FDE68A", green: "#10B981", greenBg: "#ECFDF5",
  greenBorder: "#A7F3D0", red: "#EF4444", blue: "#3B82F6",
  text: "#111827", muted: "#6B7280", white: "#FFFFFF",
};

const FONTS = {
  sans: "-apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif",
  mono: "Menlo, Monaco, 'Courier New', monospace",
};

export default function JAWSApp() {
  const [tab, setTab] = useState("setup");
  const [stops, setStops] = useState([]);
  const [sortDir, setSortDir] = useState("E→W");
  const [geocoding, setGeocoding] = useState(false);
  const [geoProgress, setGeoProgress] = useState(0);
  const [driverIdx, setDriverIdx] = useState(0);
  const [tracking, setTracking] = useState(false);
  const [driverPos, setDriverPos] = useState(null);
  const [photoForStop, setPhotoForStop] = useState(null);
  const [editGate, setEditGate] = useState({ id: null, val: "" });
  const [copied, setCopied] = useState(false);
  const [defaultCity, setDefaultCity] = useState("Miramar Beach, FL");
  const [geoDb, setGeoDb] = useState({});          // address key → {lat, lng}
  const [dbLoaded, setDbLoaded] = useState(false);
  const [editingAddr, setEditingAddr] = useState(null); // address being pinned
  const [latInput, setLatInput] = useState("");
  const [lngInput, setLngInput] = useState("");
  const [dbSaved, setDbSaved] = useState(false);
  const [publishStatus, setPublishStatus] = useState(null); // null | "publishing" | "published" | "error"
  const [sharedRoute, setSharedRoute] = useState(null);     // loaded shared route metadata
  const [checkingShared, setCheckingShared] = useState(false);

  // Publish current route to shared storage — all users of this artifact can load it
  const publishRoute = async () => {
    if (!stops.length) return;
    setPublishStatus("publishing");
    try {
      const payload = {
        stops: stops.map(s => ({
          id: s.id, address: s.address, name: s.name,
          gateCode: s.gateCode, notes: s.notes,
          geocoded: s.geocoded, status: "pending",
          timestamp: null, photo: null,
        })),
        sortDir,
        publishedAt: new Date().toLocaleString("en-US", { dateStyle: "short", timeStyle: "short" }),
        publishedBy: "Dispatch",
      };
      await fetch("/api/route", { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify(payload) });
      setPublishStatus("published");
      setTimeout(() => setPublishStatus(null), 4000);
    } catch {
      setPublishStatus("error");
      setTimeout(() => setPublishStatus(null), 3000);
    }
  };

  // Load shared route — called by the driver on their device
  const loadSharedRoute = async () => {
    setCheckingShared(true);
    try {
      const res = await fetch("/api/route");
      const data = await res.json();
      if (!data?.route) { setCheckingShared(false); return; }
      const payload = data.route;
      setStops(payload.stops);
      setSortDir(payload.sortDir || "E→W");
      setDriverIdx(0);
      setSharedRoute({ publishedAt: payload.publishedAt, count: payload.stops.length });
      setTab("driver");
    } catch {}
    setCheckingShared(false);
  };

  // Check on mount if there's a shared route waiting (from API)
  useEffect(() => {
    const check = async () => {
      try {
        const res = await fetch("/api/route");
        const data = await res.json();
        if (data?.route) {
          setSharedRoute({ publishedAt: data.route.publishedAt, count: data.route.stops.length });
        }
      } catch(e) { console.log("No shared route yet"); }
    };
    check();
  }, []);

  // ─── SEED GEO DATABASE (verified coordinates) ───────────────────────────────
  const SEED_GEO = {
    "30 dune side lane, santa rosa beach, fl 32459":                    {lat:30.341269,lng:-86.205159},
    "23 seawinds ct, santa rosa beach, fl 32459":                       {lat:30.304111,lng:-86.086802}, // converted from DMS
    "9 seawinds ct, santa rosa beach, fl 32459":                        {lat:30.303948,lng:-86.086950},
    "36 sundown ct, santa rosa beach, fl 32459":                        {lat:30.346889,lng:-86.229073},
    "111 sand dunes rd, santa rosa beach, fl 32459":                    {lat:30.338936,lng:-86.199637},
    "44 seapointe, santa rosa beach, fl 32459":                         {lat:30.306157,lng:-86.089598},
    "78 sandalwood, santa rosa beach, fl 32459":                        {lat:30.312933,lng:-86.110027},
    "33 gulf ridge, santa rosa beach, fl 32459":                        {lat:30.332145,lng:-86.176497},
    "19 daybreak ct., santa rosa beach, fl 32459":                      {lat:30.346541,lng:-86.228589},
    "14 santa clara, santa rosa beach, fl 32459":                       {lat:30.314621,lng:-86.119000},
    "384 beachfront trail, santa rosa beach, fl 32459":                 {lat:30.303727,lng:-86.086055},
    "18 hiker st., santa rosa beach, fl 32459":                         {lat:30.313238,lng:-86.099364},
    "352 cassine garden circle, santa rosa beach, fl 32459":            {lat:30.313523,lng:-86.107639},
    "106 chelsea loop, santa rosa beach, fl 32459":                     {lat:30.308320,lng:-86.099025},
    "51 north ryan st, santa rosa beach, fl 32459":                     {lat:30.309572,lng:-86.100149},
    "2305 s co. hwy 83, santa rosa beach, fl 32459":                    {lat:30.338848,lng:-86.200811},
    "28 emerald beach circle, santa rosa beach, fl 32459":              {lat:30.346927,lng:-86.202422},
    "282 emerald beach circle, santa rosa beach, fl 32459":             {lat:30.347208,lng:-86.201616},
    "118 sugar sand lane, santa rosa beach, fl 32459":                  {lat:30.313692,lng:-86.110875},
    "27 sugar sand lane, santa rosa beach, fl 32459":                   {lat:30.312375,lng:-86.110799},
    "33 sugar sand lane, santa rosa beach, fl 32459":                   {lat:30.312536,lng:-86.110799},
    "125 cabana trail, santa rosa beach, fl 32459":                     {lat:30.342450,lng:-86.200308},
    "151 shady pines dr., santa rosa beach, fl 32459":                  {lat:30.318158,lng:-86.119174},
    "83 sawgrass lane, santa rosa beach, fl 32459":                     {lat:30.313848,lng:-86.109091},
    "43 sawgrass lane, santa rosa beach, fl 32459":                     {lat:30.313624,lng:-86.109407},
    "75 emerald dunes circle, santa rosa beach, fl 32459":              {lat:30.346307,lng:-86.222140},
    "48 pointe circle, santa rosa beach, fl 32459":                     {lat:30.351436,lng:-86.239360},
    "69 canal street, santa rosa beach, fl 32459":                      {lat:30.320262,lng:-86.128686},
    "130 palmeira way, santa rosa beach, fl 32459":                     {lat:30.312606,lng:-86.109399},
    "81 palmeira way, santa rosa beach, fl 32459":                      {lat:30.312726,lng:-86.109330},
    "38 n grande beach drive, santa rosa beach, fl 32459":              {lat:30.339965,lng:-86.202961},
    "35 beachwalk lane, santa rosa beach, fl 32459":                    {lat:30.346504,lng:-86.225226},
    "64 n dothan ave, santa rosa beach, fl 32459":                      {lat:30.315989,lng:-86.121119},
    "285 gulf view circle, santa rosa beach, fl 32459":                 {lat:30.347401,lng:-86.202827},
    "90 snapper, santa rosa beach, fl 32459":                           {lat:30.344588,lng:-86.208299},
    "99 cypress passage, santa rosa beach, fl 32459":                   {lat:30.364373,lng:-86.268957},
    "238 ventana blvd, santa rosa beach, fl 32459":                     {lat:30.346286,lng:-86.202793},
    "155 baird rd, santa rosa beach, fl 32459":                         {lat:30.356148,lng:-86.257247},
    "97 hidden lake way, santa rosa beach, fl 32459":                   {lat:30.348033,lng:-86.222964},
    "104 hidden lake way, santa rosa beach, fl 32459":                  {lat:30.348033,lng:-86.222964},
    "277 hidden lake way, santa rosa beach, fl 32459":                  {lat:30.348033,lng:-86.222964},
    "354 hidden lake way, santa rosa beach, fl 32459":                  {lat:30.348033,lng:-86.222964},
    "368 hidden lake way, santa rosa beach, fl 32459":                  {lat:30.348033,lng:-86.222964},
    "378 hidden lake way, santa rosa beach, fl 32459":                  {lat:30.348033,lng:-86.222964},
    "390 hidden lake way, santa rosa beach, fl 32459":                  {lat:30.348033,lng:-86.222964},
    "416 hidden lake way, santa rosa beach, fl 32459":                  {lat:30.348033,lng:-86.222964},
    "440 hidden lake way, santa rosa beach, fl 32459":                  {lat:30.348033,lng:-86.222964},
    "468 hidden lake way, santa rosa beach, fl 32459":                  {lat:30.348033,lng:-86.222964},
    "471 hidden lake way, santa rosa beach, fl 32459":                  {lat:30.348033,lng:-86.222964},
    "12 gulf cove ct., santa rosa beach, fl 32459":                     {lat:30.314363,lng:-86.110636},
    "26 gulf cove court, santa rosa beach, fl 32459":                   {lat:30.314363,lng:-86.110636},
    "40 gulf cove, ct., santa rosa beach, fl 32459":                    {lat:30.314363,lng:-86.110636},
    "48 gulf cove ct., santa rosa beach, fl 32459":                     {lat:30.314363,lng:-86.110636},
    "50 gulf cove ct, santa rosa beach, fl 32459":                      {lat:30.314363,lng:-86.110636},
    "54 gulf cove ct., santa rosa beach, fl 32459":                     {lat:30.314363,lng:-86.110636},
    "58 gulf cove court, santa rosa beach, fl 32459":                   {lat:30.314363,lng:-86.110636},
    "62 gulf cove ct., santa rosa beach, fl 32459":                     {lat:30.314363,lng:-86.110636},
    "70 gulf cove ct, santa rosa beach, fl 32459":                      {lat:30.314363,lng:-86.110636},
    "76 gulf cove ct, santa rosa beach, fl 32459":                      {lat:30.314363,lng:-86.110636},
    "82 gulf cove ct, santa rosa beach, fl 32459":                      {lat:30.314363,lng:-86.110636},
    "4049 w county hwy 30a, santa rosa beach, fl 32459":                {lat:30.346464,lng:-86.224299},
    "4559 w co hwy 30a, santa rosa beach, fl 32459":                    {lat:30.347568,lng:-86.232577},
    "4815 w co hwy 30a, santa rosa beach, fl 32459":                    {lat:30.348885,lng:-86.236635},
    "5055 w county hwy 30a, unit 1012, santa rosa beach, fl 32459":     {lat:30.349844,lng:-86.241390},
    // Seacrest / Inlet Beach
    "99 cobia run west, seacrest, fl 32461":                            {lat:30.284342,lng:-86.020138},
    "178 seacrest beach blvd west, seacrest beach, fl 32461":           {lat:30.285304,lng:-86.022598},
    "28 pelican glide lane, inlet beach, fl 32461":                     {lat:30.288522,lng:-86.043544},
    // Miramar Beach
    "88 st francis, miramar beach, fl 32550":                           {lat:30.378839,lng:-86.379866},
    "56 st francis, miramar beach, fl 32550":                           {lat:30.380321,lng:-86.380059},
    "86 st simon cir, miramar beach, fl 32550":                         {lat:30.381728,lng:-86.380107},
    "109 st simon cir, miramar beach, fl 32550":                        {lat:30.382075,lng:-86.380091},
    "13 monaco st, miramar beach, fl 32550":                            {lat:30.375233,lng:-86.356336},
    "39 monaco, miramar beach, fl 32550":                               {lat:30.374886,lng:-86.356368},
    "57 sandy dunes cir, miramar beach, fl 32550":                      {lat:30.373271,lng:-86.346031},
    "63 sandy dunes cir, miramar beach, fl 32550":                      {lat:30.373271,lng:-86.346031},
    "96 norwood, miramar beach, fl 32550":                              {lat:30.374692,lng:-86.350820},
    "138 windrift dr, miramar beach, fl 32550":                         {lat:30.376974,lng:-86.356361},
    "5 lake ct, miramar beach, fl 32550":                               {lat:30.374610,lng:-86.354104},
    "374 rue caribe, miramar beach, fl 32550":                          {lat:30.371915,lng:-86.340955},
    "56 ballamore cove, miramar beach, fl 32550":                       {lat:30.379003,lng:-86.389741},
    "51 emerald haven, miramar beach, fl 32550":                        {lat:30.380811,lng:-86.383668},
    "37 emerald haven, miramar beach, fl 32550":                        {lat:30.380811,lng:-86.383668},
    "1495 scenic gulf, miramar beach, fl 32550":                        {lat:30.376346,lng:-86.373191},
    "1987 scenic a1 & a4, miramar beach, fl 32550":                     {lat:30.377503,lng:-86.381106},
    "1987 scenic hwy 98 unit 6, miramar beach, fl 32550":               {lat:30.377503,lng:-86.381106},
    "2841 scenic hwy 98, miramar beach, fl 32550":                      {lat:30.383035,lng:-86.429037},
    "2857 scenic hwy 98, miramar beach, fl 32550":                      {lat:30.383006,lng:-86.428905},
    "2861 scenic hwy 98, miramar beach, fl 32550":                      {lat:30.379164,lng:-86.395356},
    "2871 scenic hwy 98, miramar beach, fl 32550":                      {lat:30.379164,lng:-86.395356},
    // Destin
    "3421 scenic hwy 98, destin, fl 32541":                             {lat:30.381860,lng:-86.416372},
    "3489 scenic hwy 98, destin, fl 32541":                             {lat:30.381288,lng:-86.412199},
    "3650 scenic highway 98 unit 9 &10 destin, fl 32541":               {lat:30.379753,lng:-86.400141},
    "3650 scenic highway 98 unit 12 destin, fl 32541":                  {lat:30.379753,lng:-86.400141},
    "3650 scenic highway 98 unit 19 destin, fl 32541":                  {lat:30.379753,lng:-86.400141},
    "4510 ocean view dr, destin, fl 32541":                             {lat:30.382943,lng:-86.423246},
    "4463 luke ave #c, destin, fl 32541":                               {lat:30.386287,lng:-86.425495},
    "4579 luke ave, destin, fl 32541":                                  {lat:30.384950,lng:-86.416249},
    "4649 destiny way, destin, fl 32541":                               {lat:30.381620,lng:-86.410507},
    "4747 papaya park, destin, fl 32541":                               {lat:30.385382,lng:-86.403555},
    "4758 bonaire cay, destin, fl 32541":                               {lat:30.382809,lng:-86.401845},
    "238 kono way, destin, fl 32541":                                   {lat:30.387954,lng:-86.418911},
    "66 tarpon, destin, fl 32541":                                      {lat:30.382738,lng:-86.417153},
  };
  const watchId = useRef(null);
  const photoRef = useRef(null);
  const xlsxRef = useRef(null);

  // Aggressively normalize address for DB key matching
  // Handles: extra ", US", commas before zips, "Fl," vs "FL", trailing periods, extra spaces
  const addrKey = a => a
    .toLowerCase()
    .replace(/,?\s*\bus(a)?\b\.?$/i, "")        // strip trailing ", US" or ", USA"
    .replace(/,\s*(\d{5}(-\d{4})?)\s*$/g, " $1") // "FL, 32459" → "FL 32459"
    .replace(/\bfl\b\.?/g, "fl")                  // normalize FL/Fl/fl./FL.
    .replace(/\b(st|dr|rd|ln|ct|ave|blvd|hwy|co|n|s|e|w)\./g, "$1") // strip trailing dots in abbrevs
    .replace(/[,]+/g, " ")                         // commas → spaces
    .replace(/\s+/g, " ")                          // collapse whitespace
    .replace(/\s*(#)\s*/g, " #")                   // normalize unit #
    .trim();

  // Load geoDb from localStorage on mount — seed always base, user edits on top
  useEffect(() => {
    try {
      const norm = (k) => k.toLowerCase()
        .replace(/,?\s*\bus(a)?\b\.?$/i,"")
        .replace(/,\s*(\d{5}(-\d{4})?)\s*$/g," $1")
        .replace(/\bfl\b\.?/g,"fl")
        .replace(/\b(st|dr|rd|ln|ct|ave|blvd|hwy|co|n|s|e|w)\./g,"$1")
        .replace(/[,]+/g," ").replace(/\s+/g," ")
        .replace(/\s*(#)\s*/g," #").trim();
      const normalizedSeed = {};
      Object.entries(SEED_GEO).forEach(([k, v]) => { normalizedSeed[norm(k)] = v; });
      let userEdits = {};
      if (typeof window !== "undefined") {
        const stored = localStorage.getItem("jaws-geo-db");
        if (stored) userEdits = JSON.parse(stored);
      }
      setGeoDb({ ...normalizedSeed, ...userEdits });
    } catch(e) {
      console.error("Load error:", e);
      const normalizedSeed = {};
      const norm = (k) => k.toLowerCase().replace(/,?\s*\bus(a)?\b\.?$/i,"").replace(/,\s*(\d{5}(-\d{4})?)\s*$/g," $1").replace(/\bfl\b\.?/g,"fl").replace(/[,]+/g," ").replace(/\s+/g," ").trim();
      Object.entries(SEED_GEO).forEach(([k, v]) => { normalizedSeed[norm(k)] = v; });
      setGeoDb(normalizedSeed);
    }
    setDbLoaded(true);
  }, []);

  // Save geoDb — persist user edits to localStorage
  const saveGeoDb = (db) => {
    setGeoDb(db);
    try {
      const delta = {};
      Object.entries(db).forEach(([k, v]) => {
        const seedVal = Object.values(SEED_GEO).find(sv => 
          Math.abs(sv.lat - v.lat) < 0.000001 && Math.abs(sv.lng - v.lng) < 0.000001
        );
        if (!seedVal) delta[k] = v;
      });
      if (typeof window !== "undefined") {
        localStorage.setItem("jaws-geo-db", JSON.stringify(delta));
      }
      setDbSaved(true);
      setTimeout(() => setDbSaved(false), 2000);
    } catch(e) {
      console.error("Storage error:", e);
    }
  };

  // Resolve coordinates for an address — checks geoDb first, falls back to estimateCoords
  // Called ONCE at upload time and stored on the stop. Sort and map use stop.geocoded directly.
  const resolveCoords = useCallback((address) => {
    const key = addrKey(address);
    // Walk all geoDb keys and find the best match
    // Direct key match first
    if (geoDb[key]) return { ...geoDb[key], fromDb: true };
    // Fuzzy: try stripping unit/apt suffixes and re-matching
    const stripped = key.replace(/\s+(unit|apt|#|suite)\s*\S+\s*$/, "").trim();
    if (geoDb[stripped]) return { ...geoDb[stripped], fromDb: true };
    // Fall back to street-name estimator
    return estimateCoords(address);
  }, [geoDb]);

  // When geoDb changes (new pin saved), refresh all stop coordinates immediately
  useEffect(() => {
    if (!dbLoaded || !stops.length) return;
    setStops(prev => prev.map(s => ({
      ...s,
      geocoded: resolveCoords(s.address),
    })));
  }, [geoDb, dbLoaded]);

  const completedCount = stops.filter(s => s.status === "complete").length;
  const allDone = stops.length > 0 && completedCount === stops.length;

  // Sort uses stop.geocoded.lng directly — coordinates already resolved, no string matching here
  const sortedStops = [...stops].sort((a, b) => {
    const aLng = a.geocoded?.lng ?? 0;
    const bLng = b.geocoded?.lng ?? 0;
    return sortDir === "E→W" ? bLng - aLng : aLng - bLng;
  });

  const currentStop = sortedStops[driverIdx] || sortedStops.find(s => s.status === "pending");

  const parseExcel = (file) => {
    const reader = new FileReader();
    reader.onload = (e) => {
      const wb = XLSX.read(e.target.result, { type: "binary" });
      const ws = wb.Sheets[wb.SheetNames[0]];
      const rows = XLSX.utils.sheet_to_json(ws, { header: 1 });
      if (!rows.length) return;

      const firstRow = rows[0].map(h => String(h || "").toLowerCase().trim());
      const hasHeader = firstRow.some(h =>
        ["address", "addr", "street", "gate", "property", "name", "unit", "location"].some(k => h.includes(k))
      );

      const dataRows = hasHeader ? rows.slice(1) : rows;
      const hdrs = hasHeader ? firstRow : null;
      let aC, gC, nC, ntC;

      if (hasHeader) {
        const fc = (...kw) => hdrs.findIndex(h => kw.some(k => h.includes(k)));
        aC = fc("address", "addr", "street", "location");
        nC = fc("property", "name", "unit", "client");
        gC = fc("gate", "code", "access");
        ntC = fc("note", "comment", "special");
      } else {
        // Auto-detect address column by content — look for street/city patterns
        const addrRe = /\d+\s+\w|,\s*(fl|florida)\b|hwy|highway|blvd|boulevard|\bdrive\b|\bdr\b|\bstreet\b|\bst\b|\blane\b|\bln\b|\bcircle\b|\bcir\b|\bcourt\b|\bct\b|\bway\b|\bave\b|\broad\b|\brd\b|\bloop\b|\btrail\b|\bpath\b/i;
        const sample = dataRows.slice(0, 10).filter(r => r.length > 0);
        const numCols = Math.max(...sample.map(r => r.length));
        let bestAddrCol = 1, bestScore = -1;
        for (let c = 0; c < numCols; c++) {
          const score = sample.filter(r => addrRe.test(String(r[c] || ""))).length;
          if (score > bestScore) { bestScore = score; bestAddrCol = c; }
        }
        // Gate code: a different column with short values (numbers, #-codes, or short names)
        let bestGateCol = -1;
        for (let c = 0; c < numCols; c++) {
          if (c === bestAddrCol) continue;
          const hasContent = sample.some(r => String(r[c] || "").trim().length > 0 && String(r[c] || "").trim() !== "x");
          if (hasContent && bestGateCol === -1) bestGateCol = c;
        }
        aC = bestAddrCol;
        gC = bestGateCol;
        nC = -1;
        ntC = -1;
      }

      const parsed = dataRows
        .filter(r => r.some(c => c))
        .map((r, i) => {
          const address = String(r[aC >= 0 ? aC : 0] ?? "").trim();
          return {
            id: `s${i}-${Date.now()}`,
            address,
            name: nC >= 0 ? String(r[nC] ?? "").trim() : "",
            gateCode: gC >= 0 ? String(r[gC] ?? "").trim() : "",
            notes: ntC >= 0 ? String(r[ntC] ?? "").trim() : "",
            geocoded: address ? resolveCoords(address) : null,
            status: "pending", timestamp: null, photo: null,
          };
        })
        .filter(s => s.address);

      setStops(parsed);
      setDriverIdx(0);
    };
    reader.readAsBinaryString(file);
  };

  const loadSampleData = () => {
    const sample = [
      { address: "790 Gulf Shore Dr, Destin, FL 32541", name: "Sandcastle Villa", gateCode: "4521", notes: "Bins at back gate" },
      { address: "119 Eastern Lake Rd, Santa Rosa Beach, FL 32459", name: "Blue Heron Cottage", gateCode: "", notes: "" },
      { address: "38 S Holiday Rd, Miramar Beach, FL 32550", name: "Pelican Perch", gateCode: "7788", notes: "2 bins" },
      { address: "2000 Scenic Gulf Dr, Miramar Beach, FL 32550", name: "Sunrise Retreat", gateCode: "", notes: "Roadside pickup" },
      { address: "31 Pompano St, Santa Rosa Beach, FL 32459", name: "Aqua Bungalow", gateCode: "3311", notes: "" },
    ].map((s, i) => ({ ...s, id: `demo-${i}`, geocoded: resolveCoords(s.address), status: "pending", timestamp: null, photo: null }));
    setStops(sample);
    setDriverIdx(0);
  };

  const sortRoute = () => {
    setStops(prev => prev.map(s => ({
      ...s,
      geocoded: s.geocoded?.fromDb ? s.geocoded : getCoords(s.address),
    })));
    setGeoProgress(100);
  };

  const updateStop = useCallback((id, patch) =>
    setStops(p => p.map(s => s.id === id ? { ...s, ...patch } : s)), []);

  const markComplete = (id) => {
    const t = new Date().toLocaleTimeString("en-US", { hour: "2-digit", minute: "2-digit", hour12: true });
    updateStop(id, { status: "complete", timestamp: t });
    const idx = sortedStops.findIndex(s => s.id === id);
    const nxt = sortedStops.findIndex((s, i) => i > idx && s.status === "pending");
    if (nxt >= 0) setDriverIdx(nxt);
    else {
      const anyPending = sortedStops.findIndex(s => s.id !== id && s.status === "pending");
      if (anyPending >= 0) setDriverIdx(anyPending);
    }
  };

  const handlePhoto = (e) => {
    const file = e.target.files[0];
    if (!file || !photoForStop) return;
    const reader = new FileReader();
    reader.onload = evt => updateStop(photoForStop, { photo: evt.target.result });
    reader.readAsDataURL(file);
    e.target.value = "";
    setPhotoForStop(null);
  };

  const triggerPhoto = (id) => {
    setPhotoForStop(id);
    setTimeout(() => {
      if (photoRef.current) {
        photoRef.current.setAttribute("capture", "environment");
        photoRef.current.click();
      }
    }, 10);
  };

  const startGPS = () => {
    if (!navigator.geolocation) return;
    watchId.current = navigator.geolocation.watchPosition(
      p => setDriverPos({ lat: p.coords.latitude, lng: p.coords.longitude }),
      () => {}, { enableHighAccuracy: true }
    );
    setTracking(true);
  };

  const stopGPS = () => {
    if (watchId.current) navigator.geolocation.clearWatch(watchId.current);
    setTracking(false); setDriverPos(null);
  };

  const resetRoute = () => {
    setStops(p => p.map(s => ({ ...s, status: "pending", timestamp: null, photo: null })));
    setDriverIdx(0);
  };

  const exportCSV = () => {
    const rows = [["#", "Property", "Address", "Gate Code", "Status", "Time"]];
    sortedStops.forEach((s, i) =>
      rows.push([i + 1, s.name, s.address, s.gateCode, s.status, s.timestamp || ""])
    );
    const csv = rows.map(r => r.map(c => `"${String(c).replace(/"/g, '""')}"`).join(",")).join("\n");
    const a = Object.assign(document.createElement("a"), {
      href: URL.createObjectURL(new Blob([csv], { type: "text/csv" })),
      download: `jaws-route-${new Date().toLocaleDateString("en-US").replace(/\//g, "-")}.csv`,
    });
    a.click();
  };

  const routeDate = new Date().toLocaleDateString("en-US", { weekday: "long", year: "numeric", month: "long", day: "numeric" });

  const buildSummaryText = () => {
    const done = sortedStops.filter(s => s.status === "complete");
    const times = done.map(s => s.timestamp).filter(Boolean);
    const span = times.length >= 2 ? `${times[0]} – ${times[times.length - 1]}` : times[0] || "—";
    const lines = [
      `JAWS Services Inc. — Route Summary`,
      `Date: ${routeDate}`,
      `Stops: ${completedCount}/${stops.length} collected${times.length >= 2 ? ` (${span})` : ""}`,
      ``,
      ...sortedStops.map((s, i) => {
        const status = s.status === "complete" ? "✓" : "○";
        const time = s.timestamp ? ` — ${s.timestamp}` : "";
        const gate = s.gateCode ? ` [Gate: ${s.gateCode}]` : "";
        return `${i + 1}. ${status} ${s.name ? s.name + " — " : ""}${s.address}${gate}${time}`;
      }),
      ``,
      `Generated by JAWS Services route app`,
    ];
    return lines.join("\n");
  };

  const copySummary = () => {
    navigator.clipboard.writeText(buildSummaryText()).then(() => {
      setCopied(true);
      setTimeout(() => setCopied(false), 2500);
    });
  };

  const emailSummary = () => {
    const subject = encodeURIComponent(`JAWS Route Summary — ${new Date().toLocaleDateString("en-US")}`);
    const body = encodeURIComponent(buildSummaryText());
    window.open(`mailto:?subject=${subject}&body=${body}`);
  };

  const textSummary = () => {
    const body = encodeURIComponent(buildSummaryText());
    window.open(`sms:?&body=${body}`);
  };

  // ─── Shared style helpers ───────────────────────────────────────────────────

  const card = (ex = {}) => ({
    background: C.surface, border: `1px solid ${C.border}`, borderRadius: 10,
    padding: "14px 16px", marginBottom: 10, ...ex,
  });

  const btn = (bg = C.amber, ex = {}) => ({
    background: bg, border: "none", borderRadius: 8, padding: "10px 16px",
    fontSize: 14, fontWeight: 600, cursor: "pointer", fontFamily: FONTS.sans,
    color: bg === C.amber ? "#000" : bg === C.s2 ? C.text : "#fff",
    display: "inline-flex", alignItems: "center", gap: 6, justifyContent: "center",
    transition: "opacity 0.15s", letterSpacing: 0.3, ...ex,
  });

  const pill = (bg, color) => ({
    display: "inline-flex", alignItems: "center", padding: "3px 10px",
    borderRadius: 20, background: bg, color, fontSize: 11, fontWeight: 600,
    fontFamily: FONTS.mono, letterSpacing: 0.5,
  });

  const lbl = {
    fontSize: 11, color: C.muted, fontWeight: 600, letterSpacing: 0.8,
    textTransform: "uppercase", display: "block", marginBottom: 5, fontFamily: FONTS.sans,
  };

  // ─── SETUP TAB ──────────────────────────────────────────────────────────────

  const SetupTab = () => (
    <div style={{ padding: "16px 14px" }}>
      <label htmlFor="xlsx-input" style={{
          display: "block",
          ...card({ borderStyle: "dashed", borderWidth: 2, borderColor: C.amberBorder,
            background: C.amberBg, textAlign: "center", cursor: "pointer", padding: "28px 20px" }),
        }}
      >
        <div style={{ fontSize: 36, marginBottom: 8 }}>📋</div>
        <div style={{ fontSize: 17, fontWeight: 700, color: C.text }}>Upload route list</div>
        <div style={{ fontSize: 13, color: C.muted, marginTop: 4, lineHeight: 1.5 }}>
          Excel (.xlsx) or CSV — columns auto-detected<br />
          <span style={{ fontFamily: FONTS.mono, fontSize: 11 }}>Address · Property · Gate Code · Notes</span>
        </div>
      </label>

      {!stops.length && (
        <>
          <button onClick={loadSampleData} style={{ ...btn(C.s2, { width: "100%", border: `1px solid ${C.border}`, color: C.muted, marginTop: 0 }) }}>
            ↗ Load sample data (demo)
          </button>
          {sharedRoute && (
            <button onClick={loadSharedRoute} disabled={checkingShared} style={{ ...btn("#3B82F6", { width: "100%", color: "#fff", marginTop: 0 }) }}>
              {checkingShared ? "Loading..." : `📡 Load shared route (${sharedRoute.count} stops · ${sharedRoute.publishedAt})`}
            </button>
          )}
        </>
      )}

      {stops.length > 0 && (
        <>
          <div style={{ display: "flex", gap: 10, marginBottom: 10 }}>
            {[
              { val: stops.length, lbl: "Stops", color: C.amber },
              { val: `${stops.length}/${stops.length}`, lbl: "Geocoded", color: C.green },
              { val: completedCount, lbl: "Done", color: completedCount > 0 ? C.green : C.muted },
            ].map(stat => (
              <div key={stat.lbl} style={{ ...card({ flex: 1, textAlign: "center", padding: "10px 8px", marginBottom: 0 }) }}>
                <div style={{ fontSize: 26, fontWeight: 800, color: stat.color, lineHeight: 1, fontFamily: FONTS.mono }}>{stat.val}</div>
                <div style={{ ...lbl, marginBottom: 0, marginTop: 3, textAlign: "center" }}>{stat.lbl}</div>
              </div>
            ))}
          </div>

          <div style={card({ padding: "12px 14px" })}>
            <span style={lbl}>Route direction</span>
            <div style={{ display: "flex", gap: 8 }}>
              {["E→W", "W→E"].map(dir => (
                <button key={dir} onClick={() => setSortDir(dir)} style={{
                  ...btn(sortDir === dir ? C.amber : C.s2, { flex: 1,
                    border: `1px solid ${sortDir === dir ? C.amber : C.border}`,
                    color: sortDir === dir ? "#000" : C.muted }),
                }}>
                  {dir === "E→W" ? "🌅 East → West" : "🌇 West → East"}
                </button>
              ))}
            </div>
          </div>

          {stops.length > 0 && (
            <div style={{ ...card({ background: C.greenBg, borderColor: C.greenBorder, padding: "10px 14px", display: "flex", alignItems: "center", gap: 8 }) }}>
              <span style={{ fontSize: 16 }}>✓</span>
              <span style={{ fontSize: 13, color: C.green, fontWeight: 600 }}>
                {stops.length} stops loaded — route sorted from coordinates
              </span>
            </div>
          )}

          {/* Publish to driver */}
          {stops.length > 0 && (
            <div style={card({ padding: "14px 16px", background: "#EFF6FF", borderColor: "#BFDBFE" })}>
              <div style={{ fontWeight: 700, fontSize: 14, color: "#1E40AF", marginBottom: 4 }}>
                📡 Share with driver
              </div>
              <div style={{ fontSize: 12, color: "#3B82F6", marginBottom: 10, lineHeight: 1.5 }}>
                Publish this route so your driver can load it from their device — no upload needed on their end.
              </div>
              <button
                onClick={publishRoute}
                disabled={publishStatus === "publishing"}
                style={{ ...btn(
                  publishStatus === "published" ? C.green : "#3B82F6",
                  { width: "100%", color: "#fff", fontSize: 14, padding: "12px 16px",
                    opacity: publishStatus === "publishing" ? 0.6 : 1 }
                )}}
              >
                {publishStatus === "publishing" ? "📡 Publishing..." :
                 publishStatus === "published"  ? "✓ Route published — driver can load it now" :
                 publishStatus === "error"       ? "⚠️ Failed — try again" :
                 `📡 Publish ${stops.length} stops to driver`}
              </button>
            </div>
          )}

          <div style={{ ...lbl, marginTop: 4 }}>Loaded stops</div>
          {stops.map((stop, i) => (
            <div key={stop.id} style={{
              ...card({ padding: "10px 14px", display: "flex", alignItems: "center", gap: 12, marginBottom: 7 }),
            }}>
              <div style={{ fontFamily: FONTS.mono, fontSize: 14, fontWeight: 600, color: C.muted, minWidth: 24 }}>{i + 1}</div>
              <div style={{ flex: 1, minWidth: 0 }}>
                <div style={{ fontWeight: 600, fontSize: 14, whiteSpace: "nowrap", overflow: "hidden", textOverflow: "ellipsis" }}>
                  {stop.name || stop.address}
                </div>
                {stop.name && (
                  <div style={{ fontSize: 12, color: C.muted, whiteSpace: "nowrap", overflow: "hidden", textOverflow: "ellipsis", fontFamily: FONTS.mono }}>
                    {stop.address}
                  </div>
                )}
                {stop.geocoded?.resolvedAs && stop.geocoded.resolvedAs !== stop.address && (
                  <div style={{ fontSize: 11, color: C.green, fontFamily: FONTS.mono, marginTop: 2 }}>
                    → {stop.geocoded.resolvedAs}
                  </div>
                )}
              </div>
              <span style={pill(stop.geocoded?.fromDb ? C.greenBg : C.amberBg, stop.geocoded?.fromDb ? C.green : C.amber)}>
                  {stop.geocoded?.fromDb ? "📌 pinned" : "~ est."}
                </span>
            </div>
          ))}

          <div style={{ display: "flex", gap: 8, marginTop: 4 }}>
            <button onClick={() => { setStops([]); setDriverIdx(0); }} style={{ ...btn(C.s2, { flex: 1, border: `1px solid ${C.border}`, color: C.red }) }}>
              🗑 Clear list
            </button>
            <button onClick={resetRoute} style={{ ...btn(C.s2, { flex: 1, border: `1px solid ${C.border}`, color: C.muted }) }}>
              ↺ Reset route
            </button>
          </div>
        </>
      )}
    </div>
  );

  // ─── ROUTE TAB ──────────────────────────────────────────────────────────────

  const RouteTab = () => {
    if (!stops.length) return (
      <div style={{ textAlign: "center", color: C.muted, paddingTop: 60, padding: "60px 20px 20px" }}>
        <div style={{ fontSize: 40 }}>🗺️</div>
        <div style={{ fontWeight: 600, fontSize: 16, marginTop: 12 }}>No stops loaded</div>
        <div style={{ fontSize: 13, marginTop: 4 }}>Go to Setup to upload a route list</div>
      </div>
    );

    return (
      <div style={{ padding: "14px 14px" }}>
        {stops.some(s => !s.geocoded) && (
          <div style={{ ...card({ background: C.amberBg, borderColor: C.amberBorder, display: "flex", gap: 10, alignItems: "flex-start" }) }}>
            <span style={{ fontSize: 18 }}>⚠️</span>
            <span style={{ fontSize: 13, color: "#92400E" }}>
              {stops.filter(s => !s.geocoded).length} stop(s) not geocoded — route order may be inaccurate. Run geocoding in Setup.
            </span>
          </div>
        )}

        <div style={{ display: "flex", gap: 8, marginBottom: 12 }}>
          {["E→W", "W→E"].map(dir => (
            <button key={dir} onClick={() => setSortDir(dir)} style={{
              ...btn(sortDir === dir ? C.amber : C.s2, { flex: 1,
                border: `1px solid ${sortDir === dir ? C.amber : C.border}`,
                color: sortDir === dir ? "#000" : C.muted, fontSize: 13 }),
            }}>
              {dir === "E→W" ? "🌅 E→W" : "🌇 W→E"}
            </button>
          ))}
          <button onClick={exportCSV} style={{ ...btn(C.s2, { flex: 1, border: `1px solid ${C.border}`, color: C.blue, fontSize: 13 }) }}>
            ↓ Export CSV
          </button>
        </div>

        {sortedStops.map((stop, i) => (
          <div key={stop.id} style={{
            ...card({ borderLeft: `4px solid ${stop.status === "complete" ? C.green : C.amber}`, opacity: stop.status === "complete" ? 0.6 : 1, padding: "12px 14px" }),
          }}>
            <div style={{ display: "flex", alignItems: "flex-start", gap: 12 }}>
              <div style={{ fontFamily: FONTS.mono, fontSize: 18, fontWeight: 700, color: stop.status === "complete" ? C.green : C.amber, minWidth: 28, lineHeight: 1 }}>
                {stop.status === "complete" ? "✓" : i + 1}
              </div>
              <div style={{ flex: 1, minWidth: 0 }}>
                {stop.name && <div style={{ fontWeight: 600, fontSize: 15, marginBottom: 2 }}>{stop.name}</div>}

                {/* Address row — left side opens Geo DB, right side opens Google Maps */}
                <div style={{ display: "flex", alignItems: "center", gap: 6, marginBottom: 10 }}>
                  <button
                    onClick={() => {
                      const stored = geoDb[addrKey(stop.address)];
                      setEditingAddr(stop.address);
                      setLatInput(stored ? stored.lat.toFixed(6) : "");
                      setLngInput(stored ? stored.lng.toFixed(6) : "");
                      setTab("geodb");
                    }}
                    style={{
                      flex: 1, textAlign: "left", background: "none", border: "none",
                      padding: 0, cursor: "pointer", minWidth: 0,
                    }}
                  >
                    <span style={{
                      fontFamily: FONTS.mono, color: geoDb[addrKey(stop.address)] ? C.green : C.blue,
                      fontSize: 12, textDecoration: "none", display: "flex", alignItems: "center", gap: 4,
                    }}>
                      {geoDb[addrKey(stop.address)] ? "📌" : "📍"}
                      <span style={{ whiteSpace: "nowrap", overflow: "hidden", textOverflow: "ellipsis", display: "block" }}>
                        {stop.address}
                      </span>
                    </span>
                    <span style={{ fontSize: 10, color: geoDb[addrKey(stop.address)] ? C.green : C.muted, fontFamily: FONTS.mono }}>
                      {geoDb[addrKey(stop.address)]
                        ? `✓ pinned ${geoDb[addrKey(stop.address)].lat.toFixed(4)}, ${geoDb[addrKey(stop.address)].lng.toFixed(4)}`
                        : "tap to fix pin location →"}
                    </span>
                  </button>
                  <a
                    href={navUrl(stop)}
                    target="_blank" rel="noopener noreferrer"
                    style={{ ...btn(C.s2, { padding: "5px 8px", fontSize: 11, border: `1px solid ${C.border}`, color: C.blue, textDecoration: "none", flexShrink: 0 }) }}
                  >
                    🗺️
                  </a>
                </div>

                <span style={lbl}>Gate code</span>
                {editGate.id === stop.id ? (
                  <div style={{ display: "flex", gap: 6 }}>
                    <input
                      value={editGate.val}
                      onChange={e => setEditGate(g => ({ ...g, val: e.target.value }))}
                      onKeyDown={e => { if (e.key === "Enter") { updateStop(stop.id, { gateCode: editGate.val }); setEditGate({ id: null, val: "" }); } }}
                      autoFocus
                      style={{ flex: 1, background: C.s2, border: `1px solid ${C.border}`, borderRadius: 6, padding: "7px 10px", fontSize: 16, fontFamily: FONTS.mono, letterSpacing: 4, color: C.text, outline: "none" }}
                      placeholder="Enter code..."
                    />
                    <button onClick={() => { updateStop(stop.id, { gateCode: editGate.val }); setEditGate({ id: null, val: "" }); }} style={btn(C.green, { padding: "7px 12px" })}>✓</button>
                    <button onClick={() => setEditGate({ id: null, val: "" })} style={{ ...btn(C.s2, { padding: "7px 12px", border: `1px solid ${C.border}` }) }}>✕</button>
                  </div>
                ) : (
                  <div onClick={() => setEditGate({ id: stop.id, val: stop.gateCode || "" })}
                    style={{ display: "inline-flex", alignItems: "center", gap: 8, cursor: "pointer",
                      background: C.s2, borderRadius: 6, padding: "7px 12px",
                      border: `1px solid ${stop.gateCode ? C.amberBorder : C.border}` }}>
                    <span style={{ fontFamily: FONTS.mono, fontSize: 18, color: stop.gateCode ? C.amber : C.muted, letterSpacing: 4 }}>
                      {stop.gateCode || "— tap to add —"}
                    </span>
                    <span style={{ fontSize: 12, color: C.muted }}>✎</span>
                  </div>
                )}

                {stop.notes && <div style={{ marginTop: 8, fontSize: 13, color: C.muted }}>📝 {stop.notes}</div>}
                {stop.status === "complete" && stop.timestamp && (
                  <span style={{ ...pill(C.greenBg, C.green), marginTop: 8, display: "inline-flex" }}>✓ collected {stop.timestamp}</span>
                )}
              </div>
            </div>
          </div>
        ))}
      </div>
    );
  };

  // ─── DRIVER TAB ─────────────────────────────────────────────────────────────

  const DriverTab = () => {
    if (!stops.length) return (
      <div style={{ textAlign: "center", color: C.muted, padding: "60px 20px 20px" }}>
        <div style={{ fontSize: 40 }}>🚛</div>
        <div style={{ fontWeight: 600, fontSize: 16, marginTop: 12 }}>No route loaded</div>
        <div style={{ fontSize: 13, marginTop: 4 }}>Upload a route list in the Setup tab</div>
      </div>
    );

    const curIdx = sortedStops.indexOf(currentStop);
    const nextStop = sortedStops.find((s, i) => i > curIdx && s.status === "pending");

    return (
      <div style={{ padding: "14px 14px" }}>
        {/* GPS + progress row */}
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 12 }}>
          <div style={{ display: "flex", alignItems: "center", gap: 7, fontSize: 13, color: tracking ? C.green : C.muted }}>
            <div style={{
              width: 8, height: 8, borderRadius: "50%",
              background: tracking ? C.green : C.muted,
              boxShadow: tracking ? `0 0 0 2px ${C.greenBg}` : "none",
            }} />
            {tracking ? (driverPos ? `${driverPos.lat.toFixed(4)}, ${driverPos.lng.toFixed(4)}` : "Acquiring GPS...") : "GPS off"}
          </div>
          <button onClick={tracking ? stopGPS : startGPS} style={btn(tracking ? "#FEF2F2" : C.greenBg, {
            color: tracking ? C.red : C.green, border: `1px solid ${tracking ? "#FECACA" : C.greenBorder}`,
            padding: "6px 12px", fontSize: 12 })}>
            {tracking ? "⏹ Stop GPS" : "▶ Start GPS"}
          </button>
        </div>

        <div style={{ marginBottom: 16 }}>
          <div style={{ display: "flex", justifyContent: "space-between", marginBottom: 6 }}>
            <span style={lbl}>Route progress</span>
            <span style={{ fontFamily: FONTS.mono, fontSize: 12, fontWeight: 600, color: allDone ? C.green : C.amber }}>
              {completedCount} / {stops.length}
            </span>
          </div>
          <div style={{ background: C.s2, borderRadius: 4, height: 8, overflow: "hidden" }}>
            <div style={{
              background: allDone ? C.green : C.amber, height: "100%",
              width: `${(completedCount / stops.length) * 100}%`,
              transition: "width 0.5s ease", borderRadius: 4,
            }} />
          </div>
        </div>

        {allDone ? (
          <div style={{ ...card({ textAlign: "center", padding: "36px 20px", background: C.greenBg, borderColor: C.greenBorder }) }}>
            <div style={{ fontSize: 56 }}>🎉</div>
            <div style={{ fontSize: 26, fontWeight: 800, color: C.green, marginBottom: 6 }}>Route complete!</div>
            <div style={{ fontSize: 14, color: C.muted }}>All {stops.length} stops collected</div>
            <div style={{ display: "flex", gap: 10, marginTop: 20 }}>
              <button onClick={exportCSV} style={{ ...btn(C.green, { flex: 1, color: "#fff" }) }}>↓ Export summary</button>
              <button onClick={resetRoute} style={{ ...btn(C.s2, { flex: 1, border: `1px solid ${C.border}`, color: C.muted }) }}>↺ Reset</button>
            </div>
          </div>
        ) : currentStop ? (
          <>
            {/* CURRENT STOP — BIG CARD */}
            <div style={{ ...card({ border: `2px solid ${C.amber}`, background: C.amberBg, padding: 18, marginBottom: 12 }) }}>
              <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 12 }}>
                <span style={pill(C.amber + "33", "#92400E")}>Stop {sortedStops.indexOf(currentStop) + 1} of {stops.length}</span>
                <span style={{ fontFamily: FONTS.mono, fontSize: 11, color: C.muted }}>{stops.length - completedCount} remaining</span>
              </div>

              {currentStop.name && (
                <div style={{ fontSize: 22, fontWeight: 800, lineHeight: 1.2, marginBottom: 6, color: C.text }}>{currentStop.name}</div>
              )}

              <a href={navUrl(currentStop)}
                target="_blank" rel="noopener noreferrer"
                style={{ fontFamily: FONTS.mono, color: C.blue, fontSize: 13, textDecoration: "none", display: "block", marginBottom: 16, lineHeight: 1.4 }}>
                📍 {currentStop.address} ↗
              </a>

              {/* GATE CODE — prominent */}
              {currentStop.gateCode ? (
                <div style={{
                  background: C.white, border: `2px solid ${C.amber}`, borderRadius: 10,
                  padding: "14px 18px", marginBottom: 14, display: "flex", alignItems: "center", gap: 14,
                }}>
                  <span style={{ fontSize: 26 }}>🔐</span>
                  <div>
                    <div style={{ ...lbl, color: "#92400E", marginBottom: 2 }}>Gate code</div>
                    <div style={{ fontFamily: FONTS.mono, fontSize: 30, fontWeight: 700, color: C.amber, letterSpacing: 8 }}>
                      {currentStop.gateCode}
                    </div>
                  </div>
                </div>
              ) : (
                <div style={{ ...card({ display: "flex", alignItems: "center", gap: 8, padding: "10px 14px", marginBottom: 12, background: C.white }) }}>
                  <span style={{ fontSize: 18 }}>🔓</span>
                  <span style={{ fontSize: 13, color: C.muted }}>No gate code — open access</span>
                </div>
              )}

              {currentStop.notes && (
                <div style={{ background: C.white, borderRadius: 8, padding: "10px 14px", marginBottom: 14, fontSize: 14, lineHeight: 1.4 }}>
                  📝 {currentStop.notes}
                </div>
              )}

              {currentStop.photo && (
                <div style={{ marginBottom: 14, position: "relative" }}>
                  <img src={currentStop.photo} alt="Collection"
                    style={{ width: "100%", maxHeight: 180, objectFit: "cover", borderRadius: 8, display: "block" }} />
                  <span style={{ ...pill(C.greenBg, C.green), position: "absolute", top: 8, right: 8, background: C.green, color: "#fff" }}>
                    📷 saved
                  </span>
                </div>
              )}

              <div style={{ display: "flex", gap: 10, marginBottom: 10 }}>
                <button onClick={() => triggerPhoto(currentStop.id)}
                  style={{ ...btn(C.white, { flex: 1, border: `1px solid ${C.border}`, color: C.text }) }}>
                  📷 {currentStop.photo ? "Retake photo" : "Take photo"}
                </button>
                <a href={navUrl(currentStop)}
                  target="_blank" rel="noopener noreferrer"
                  style={{ ...btn(C.blue, { flex: 1, color: "#fff", textDecoration: "none" }) }}>
                  🗺️ Navigate
                </a>
              </div>

              <button onClick={() => markComplete(currentStop.id)}
                style={{ ...btn(C.green, { width: "100%", padding: "16px 16px", fontSize: 18, color: "#fff", fontWeight: 700 }) }}>
                ✓ Mark collected
              </button>
            </div>

            {/* NEXT STOP PREVIEW */}
            {nextStop && (
              <div style={{ ...card({ padding: "12px 14px", opacity: 0.75, marginBottom: 16 }) }}>
                <span style={lbl}>Next stop</span>
                <div style={{ fontWeight: 600, fontSize: 16 }}>{nextStop.name || nextStop.address}</div>
                {nextStop.name && <div style={{ fontFamily: FONTS.mono, fontSize: 12, color: C.muted }}>{nextStop.address}</div>}
                {nextStop.gateCode && <div style={{ color: C.amber, fontSize: 13, marginTop: 5, fontFamily: FONTS.mono }}>🔐 {nextStop.gateCode}</div>}
              </div>
            )}
          </>
        ) : null}

        {/* ALL STOPS LIST */}
        <span style={{ ...lbl, marginTop: 4 }}>All stops — tap to jump</span>
        {sortedStops.map((stop, i) => {
          const isCurrent = stop.id === currentStop?.id;
          return (
            <div key={stop.id} onClick={() => setDriverIdx(i)} style={{
              display: "flex", alignItems: "center", padding: "10px 12px",
              borderRadius: 8, marginBottom: 6, cursor: "pointer",
              background: isCurrent ? C.amberBg : C.surface,
              border: `1px solid ${isCurrent ? C.amber : C.border}`,
              opacity: stop.status === "complete" ? 0.45 : 1, transition: "all 0.15s",
            }}>
              <div style={{
                fontFamily: FONTS.mono, fontSize: 15, fontWeight: 700, minWidth: 26,
                color: stop.status === "complete" ? C.green : isCurrent ? C.amber : C.muted,
              }}>
                {stop.status === "complete" ? "✓" : i + 1}
              </div>
              <div style={{ flex: 1, minWidth: 0 }}>
                <div style={{ fontSize: 14, fontWeight: 600, whiteSpace: "nowrap", overflow: "hidden", textOverflow: "ellipsis" }}>
                  {stop.name || stop.address}
                </div>
                {stop.status === "complete" && stop.timestamp && (
                  <div style={{ fontFamily: FONTS.mono, fontSize: 11, color: C.green }}>{stop.timestamp}</div>
                )}
              </div>
              <div style={{ display: "flex", gap: 4 }}>
                {stop.gateCode && <span style={pill(C.amberBg, "#92400E")}>🔐</span>}
                {stop.photo && <span style={pill(C.greenBg, C.green)}>📷</span>}
              </div>
            </div>
          );
        })}
      </div>
    );
  };

  // ─── SUMMARY TAB ────────────────────────────────────────────────────────────

  const SummaryTab = () => {
    const done = sortedStops.filter(s => s.status === "complete");
    const pending = sortedStops.filter(s => s.status === "pending");
    const times = done.map(s => s.timestamp).filter(Boolean);
    const timeSpan = times.length >= 2 ? `${times[0]} – ${times[times.length - 1]}` : times[0] || null;
    const photosCount = done.filter(s => s.photo).length;

    if (!stops.length) return (
      <div style={{ textAlign: "center", color: C.muted, padding: "60px 20px 20px" }}>
        <div style={{ fontSize: 40 }}>📊</div>
        <div style={{ fontWeight: 600, fontSize: 16, marginTop: 12 }}>No route loaded</div>
        <div style={{ fontSize: 13, marginTop: 4 }}>Upload a route list in the Setup tab</div>
      </div>
    );

    return (
      <div style={{ padding: "14px 14px" }}>

        {/* Date + share strip */}
        <div style={card({ padding: "14px 16px", marginBottom: 10 })}>
          <div style={{ fontSize: 15, fontWeight: 700, color: C.text, marginBottom: 2 }}>
            {routeDate}
          </div>
          <div style={{ fontFamily: FONTS.mono, fontSize: 12, color: C.muted }}>
            JAWS Services Inc. — Route Summary
          </div>
        </div>

        {/* Stats row */}
        <div style={{ display: "flex", gap: 10, marginBottom: 10 }}>
          {[
            { val: completedCount, total: stops.length, lbl: "Collected", color: completedCount === stops.length ? C.green : C.amber },
            { val: pending.length, total: null, lbl: "Pending", color: pending.length > 0 ? C.amber : C.muted },
            { val: photosCount, total: null, lbl: "Photos", color: photosCount > 0 ? C.blue : C.muted },
          ].map(stat => (
            <div key={stat.lbl} style={{ ...card({ flex: 1, textAlign: "center", padding: "10px 8px", marginBottom: 0 }) }}>
              <div style={{ fontSize: 22, fontWeight: 800, color: stat.color, lineHeight: 1, fontFamily: FONTS.mono }}>
                {stat.val}{stat.total !== null ? <span style={{ fontSize: 14, color: C.muted }}>/{stat.total}</span> : ""}
              </div>
              <div style={{ ...lbl, marginBottom: 0, marginTop: 3, textAlign: "center" }}>{stat.lbl}</div>
            </div>
          ))}
        </div>

        {timeSpan && (
          <div style={card({ padding: "10px 16px", display: "flex", alignItems: "center", gap: 10 })}>
            <span style={{ fontSize: 18 }}>⏱️</span>
            <div>
              <div style={{ ...lbl, marginBottom: 1 }}>Time span</div>
              <div style={{ fontFamily: FONTS.mono, fontSize: 15, fontWeight: 600, color: C.text }}>{timeSpan}</div>
            </div>
          </div>
        )}

        {/* Share buttons */}
        <div style={{ display: "flex", gap: 8, marginBottom: 14 }}>
          <button onClick={copySummary} style={{
            ...btn(copied ? C.green : C.s2, { flex: 1, border: `1px solid ${copied ? C.greenBorder : C.border}`,
              color: copied ? C.green : C.text, fontSize: 13 }),
          }}>
            {copied ? "✓ Copied!" : "📋 Copy"}
          </button>
          <button onClick={emailSummary} style={{ ...btn(C.s2, { flex: 1, border: `1px solid ${C.border}`, color: C.blue, fontSize: 13 }) }}>
            ✉️ Email
          </button>
          <button onClick={textSummary} style={{ ...btn(C.s2, { flex: 1, border: `1px solid ${C.border}`, color: C.green, fontSize: 13 }) }}>
            💬 Text
          </button>
          <button onClick={exportCSV} style={{ ...btn(C.s2, { flex: 1, border: `1px solid ${C.border}`, color: C.muted, fontSize: 13 }) }}>
            ↓ CSV
          </button>
        </div>

        {/* Stop-by-stop list */}
        <span style={lbl}>Stop details</span>
        {sortedStops.map((stop, i) => {
          const isDone = stop.status === "complete";
          return (
            <div key={stop.id} style={{
              ...card({
                borderLeft: `4px solid ${isDone ? C.green : C.border}`,
                opacity: isDone ? 1 : 0.5,
                padding: "12px 14px",
                marginBottom: 8,
              }),
            }}>
              <div style={{ display: "flex", gap: 12 }}>
                {/* Photo thumbnail or placeholder */}
                <div style={{
                  width: 56, height: 56, borderRadius: 8, flexShrink: 0,
                  overflow: "hidden", background: C.s2, border: `1px solid ${C.border}`,
                  display: "flex", alignItems: "center", justifyContent: "center",
                }}>
                  {stop.photo
                    ? <img src={stop.photo} alt="" style={{ width: "100%", height: "100%", objectFit: "cover" }} />
                    : <span style={{ fontSize: 22, opacity: 0.3 }}>📷</span>
                  }
                </div>

                <div style={{ flex: 1, minWidth: 0 }}>
                  <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", marginBottom: 2 }}>
                    <span style={{ fontFamily: FONTS.mono, fontSize: 12, color: C.muted }}>#{i + 1}</span>
                    {isDone
                      ? <span style={pill(C.greenBg, C.green)}>✓ {stop.timestamp}</span>
                      : <span style={pill(C.s2, C.muted)}>pending</span>
                    }
                  </div>
                  {stop.name && (
                    <div style={{ fontWeight: 700, fontSize: 14, lineHeight: 1.2, marginBottom: 2 }}>{stop.name}</div>
                  )}
                  <div style={{ fontFamily: FONTS.mono, fontSize: 12, color: C.muted, whiteSpace: "nowrap", overflow: "hidden", textOverflow: "ellipsis" }}>
                    {stop.address}
                  </div>
                  {stop.gateCode && (
                    <div style={{ marginTop: 4, display: "flex", alignItems: "center", gap: 6 }}>
                      <span style={{ fontSize: 12 }}>🔐</span>
                      <span style={{ fontFamily: FONTS.mono, fontSize: 13, fontWeight: 600, color: C.amber, letterSpacing: 3 }}>{stop.gateCode}</span>
                    </div>
                  )}
                  {stop.notes && (
                    <div style={{ fontSize: 12, color: C.muted, marginTop: 4 }}>📝 {stop.notes}</div>
                  )}
                </div>
              </div>

              {/* Full-size photo expand on tap */}
              {stop.photo && (
                <div style={{ marginTop: 10 }}>
                  <img src={stop.photo} alt="Collection photo"
                    style={{ width: "100%", borderRadius: 8, maxHeight: 160, objectFit: "cover", display: "block" }} />
                </div>
              )}
            </div>
          );
        })}

        {completedCount > 0 && (
          <div style={{ ...card({ background: C.greenBg, borderColor: C.greenBorder, textAlign: "center", padding: "16px 14px" }) }}>
            <div style={{ fontSize: 13, color: C.green, fontWeight: 600 }}>
              {completedCount === stops.length
                ? `🎉 Full route complete — ${completedCount} stops collected`
                : `${completedCount} of ${stops.length} stops collected so far`}
            </div>
            <div style={{ marginTop: 10, display: "flex", gap: 8, justifyContent: "center" }}>
              <button onClick={copySummary} style={{ ...btn(C.green, { color: "#fff", fontSize: 13 }) }}>
                {copied ? "✓ Copied!" : "📋 Copy summary"}
              </button>
              <button onClick={emailSummary} style={{ ...btn(C.s2, { border: `1px solid ${C.border}`, color: C.blue, fontSize: 13 }) }}>
                ✉️ Email
              </button>
            </div>
          </div>
        )}
      </div>
    );
  };

  // ─── MAP TAB ────────────────────────────────────────────────────────────────

  const [mapSelectedId, setMapSelectedId] = useState(null);

  const navUrl = (stop) => {
    const c = stop.geocoded;
    if (c?.lat && c?.lng) return `https://www.google.com/maps/dir/?api=1&destination=${c.lat},${c.lng}&travelmode=driving`;
    return `https://www.google.com/maps/dir/?api=1&destination=${encodeURIComponent(stop.address)}&travelmode=driving`;
  };

  const MapTab = () => {
    const pts = sortedStops.filter(s => s.geocoded?.lat);
    const selected = sortedStops.find(s => s.id === mapSelectedId);

    if (!stops.length) return (
      <div style={{ textAlign: "center", color: C.muted, padding: "60px 20px" }}>
        <div style={{ fontSize: 40 }}>📍</div>
        <div style={{ fontWeight: 600, fontSize: 16, marginTop: 12 }}>No route loaded</div>
        <div style={{ fontSize: 13, marginTop: 4 }}>Upload a list in Setup first</div>
      </div>
    );

    // Google Maps full route URL using lat/lng waypoints
    const buildRouteUrl = () => {
      if (!pts.length) return null;
      if (pts.length === 1) return `https://maps.google.com/?q=${pts[0].geocoded.lat},${pts[0].geocoded.lng}`;
      const chunks = [];
      for (let i = 0; i < pts.length; i += 9) chunks.push(pts.slice(i, i + 9));
      const c = chunks[0];
      const origin = `${c[0].geocoded.lat},${c[0].geocoded.lng}`;
      const dest   = `${c[c.length-1].geocoded.lat},${c[c.length-1].geocoded.lng}`;
      const wps    = c.slice(1,-1).map(s=>`${s.geocoded.lat},${s.geocoded.lng}`).join("/");
      return `https://www.google.com/maps/dir/${origin}/${wps ? wps + "/" : ""}${dest}`;
    };
    const fullRouteUrl = buildRouteUrl();

    // SVG route strip — plots stops east-to-west using real lng values
    const RouteStrip = () => {
      if (pts.length < 2) return null;
      const W = 340, H = 72, PAD = 18;
      const lngs = pts.map(s => s.geocoded.lng);
      const minL = Math.min(...lngs), maxL = Math.max(...lngs);
      const toX = lng => PAD + ((lng - minL) / (maxL - minL || 1)) * (W - PAD*2);
      const midY = H / 2;
      const pathD = pts.map((s,i) => `${i===0?"M":"L"}${toX(s.geocoded.lng)},${midY}`).join(" ");
      return (
        <div style={{ padding: "10px 14px 0" }}>
          <svg viewBox={`0 0 ${W} ${H}`} style={{ width:"100%", borderRadius:8, background:"#F0F4FF", border:`1px solid ${C.border}` }}>
            {/* Road line */}
            <line x1={PAD} y1={midY} x2={W-PAD} y2={midY} stroke="#D1D5DB" strokeWidth="4" strokeLinecap="round"/>
            {/* Route path */}
            <path d={pathD} stroke="#F59E0B" strokeWidth="2" strokeDasharray="5 4" fill="none"/>
            {/* Stop dots */}
            {pts.map((s, i) => {
              const x = toX(s.geocoded.lng);
              const isDone = s.status === "complete";
              const isCur  = s.id === currentStop?.id;
              const isSel  = s.id === mapSelectedId;
              const r = isCur || isSel ? 10 : 7;
              const fill = isDone ? "#10B981" : isCur ? "#F59E0B" : "#6B7280";
              return (
                <g key={s.id} onClick={() => setMapSelectedId(p => p === s.id ? null : s.id)} style={{cursor:"pointer"}}>
                  {(isCur||isSel) && <circle cx={x} cy={midY} r={r+5} fill={fill} opacity="0.2"/>}
                  <circle cx={x} cy={midY} r={r} fill={fill} stroke="white" strokeWidth="1.5"/>
                  <text x={x} y={midY+0.5} textAnchor="middle" dominantBaseline="middle"
                    fontSize={isCur||isSel?7:6} fontWeight="700" fill={isDone?"white":isCur?"black":"white"} fontFamily="monospace">
                    {isDone ? "✓" : i+1}
                  </text>
                </g>
              );
            })}
            {/* E / W labels */}
            <text x={4} y={midY+3} fontSize="9" fill="#9CA3AF" fontFamily="sans-serif">W</text>
            <text x={W-4} y={midY+3} fontSize="9" fill="#9CA3AF" fontFamily="sans-serif" textAnchor="end">E</text>
          </svg>
          <div style={{ fontSize: 10, color: C.muted, textAlign:"center", marginTop:4 }}>
            Tap a dot to preview · sorted {sortDir}
          </div>
        </div>
      );
    };

    return (
      <div style={{ paddingBottom: 20 }}>
        {/* Header */}
        <div style={{ padding:"10px 14px", background:C.amberBg, borderBottom:`1px solid ${C.amberBorder}`, display:"flex", alignItems:"center", justifyContent:"space-between" }}>
          <div style={{ fontSize:12, color:"#92400E", fontWeight:600 }}>
            {pts.length} stops · tap stop to preview
          </div>
          {fullRouteUrl && (
            <a href={fullRouteUrl} target="_blank" rel="noopener noreferrer"
              style={{ ...btn(C.amber, { padding:"6px 12px", fontSize:12, textDecoration:"none" }) }}>
              🗺️ Full route in Maps
            </a>
          )}
        </div>

        {/* Route strip SVG */}
        <RouteStrip />

        {/* Selected stop — Google Maps iframe using lat/lng */}
        {selected?.geocoded?.lat && (
          <div style={{ padding:"12px 14px 0" }}>
            <div style={{ borderRadius:10, overflow:"hidden", border:`1px solid ${C.border}`, marginBottom:10 }}>
              <iframe
                key={selected.id}
                src={`https://maps.google.com/maps?q=${selected.geocoded.lat},${selected.geocoded.lng}&z=17&output=embed`}
                width="100%" height="220"
                style={{ border:"none", display:"block" }}
                allowFullScreen loading="lazy"
              />
            </div>
            <div style={card({ padding:"10px 14px" })}>
              <div style={{ display:"flex", justifyContent:"space-between", alignItems:"flex-start", gap:10 }}>
                <div style={{ flex:1, minWidth:0 }}>
                  <div style={{ fontFamily:FONTS.mono, fontSize:10, color:C.muted }}>
                    Stop #{sortedStops.indexOf(selected)+1} · {selected.geocoded.lat.toFixed(5)}, {selected.geocoded.lng.toFixed(5)}
                  </div>
                  {selected.name && <div style={{ fontWeight:700, fontSize:15, marginTop:2 }}>{selected.name}</div>}
                  <div style={{ fontSize:12, color:C.muted, marginTop:1 }}>{selected.address}</div>
                  {selected.gateCode && (
                    <div style={{ color:C.amber, fontSize:15, marginTop:4, fontFamily:FONTS.mono, letterSpacing:3, fontWeight:700 }}>
                      🔐 {selected.gateCode}
                    </div>
                  )}
                </div>
                <div style={{ display:"flex", flexDirection:"column", gap:6, flexShrink:0 }}>
                  <a href={navUrl(selected)} target="_blank" rel="noopener noreferrer"
                    style={{ ...btn(C.blue, { padding:"7px 12px", fontSize:12, color:"#fff", textDecoration:"none" }) }}>
                    🧭 Navigate
                  </a>
                  <button onClick={() => setMapSelectedId(null)}
                    style={{ ...btn(C.s2, { padding:"5px 10px", fontSize:12, border:`1px solid ${C.border}`, color:C.muted }) }}>
                    ✕ Close
                  </button>
                </div>
              </div>
            </div>
          </div>
        )}

        {/* Stop list */}
        <div style={{ padding:"12px 14px 0" }}>
          <span style={lbl}>{pts.length} stops — tap to preview on map</span>
          {sortedStops.map((stop, i) => {
            const isSelected = stop.id === mapSelectedId;
            const isCurrent  = stop.id === currentStop?.id;
            const isDone     = stop.status === "complete";
            return (
              <div key={stop.id}
                onClick={() => setMapSelectedId(prev => prev === stop.id ? null : stop.id)}
                style={{
                  display:"flex", alignItems:"center", padding:"9px 12px",
                  borderRadius:8, marginBottom:6, cursor:"pointer",
                  background: isSelected ? C.amberBg : C.surface,
                  border:`1px solid ${isSelected ? C.amber : C.border}`,
                  opacity: isDone ? 0.5 : 1,
                }}>
                <div style={{ fontFamily:FONTS.mono, fontSize:13, fontWeight:700, minWidth:28,
                  color: isDone ? C.green : isCurrent ? C.amber : C.muted }}>
                  {isDone ? "✓" : i+1}
                </div>
                <div style={{ flex:1, minWidth:0 }}>
                  <div style={{ fontSize:13, fontWeight:600, whiteSpace:"nowrap", overflow:"hidden", textOverflow:"ellipsis" }}>
                    {stop.name || stop.address}
                  </div>
                  {stop.geocoded?.lat && (
                    <div style={{ fontFamily:FONTS.mono, fontSize:10, color:C.muted }}>
                      {stop.geocoded.lat.toFixed(4)}, {stop.geocoded.lng.toFixed(4)}
                      {stop.geocoded.fromDb ? " 📌" : " ~"}
                    </div>
                  )}
                </div>
                <a href={navUrl(stop)} target="_blank" rel="noopener noreferrer"
                  onClick={e => e.stopPropagation()}
                  style={{ ...btn(C.s2, { padding:"5px 8px", fontSize:11, border:`1px solid ${C.border}`, color:C.blue, textDecoration:"none", flexShrink:0, marginLeft:6 }) }}>
                  🧭
                </a>
              </div>
            );
          })}
        </div>
      </div>
    );
  };

  // ─── GEO DB TAB ─────────────────────────────────────────────────────────────

  // Initialize the DB picker map when editingAddr changes
  useEffect(() => {
    if (!editingAddr) return;

    const initMap = async () => {
      // Nothing to init — using iframe embed instead
    };

    setTimeout(initMap, 50);
    return () => {
      if (dbLeaflet.current) { dbLeaflet.current.remove(); dbLeaflet.current = null; }
    };
  }, [editingAddr]);

  const saveCoord = () => {
    const lat = parseFloat(latInput);
    const lng = parseFloat(lngInput);
    if (isNaN(lat) || isNaN(lng)) return;
    const newDb = { ...geoDb, [addrKey(editingAddr)]: { lat, lng } };
    saveGeoDb(newDb);
    // Patch matching stops directly with the new lat/lng — no string re-lookup needed
    setStops(prev => prev.map(s =>
      addrKey(s.address) === addrKey(editingAddr)
        ? { ...s, geocoded: { lat, lng, fromDb: true } }
        : s
    ));
    setEditingAddr(null);
    setLatInput(""); setLngInput("");
  };

  const deleteCoord = (addr) => {
    const newDb = { ...geoDb };
    delete newDb[addrKey(addr)];
    saveGeoDb(newDb);
  };

  const geoXlsxRef = useRef(null);

  const importGeoSpreadsheet = (file) => {
    const reader = new FileReader();
    reader.onload = (e) => {
      const wb = XLSX.read(e.target.result, { type: "binary" });
      const ws = wb.Sheets[wb.SheetNames[0]];
      const rows = XLSX.utils.sheet_to_json(ws, { header: 1 });
      if (!rows.length) return;
      const hdrs = rows[0].map(h => String(h || "").toLowerCase().trim());
      const aC = hdrs.findIndex(h => h.includes("address"));
      const latC = hdrs.findIndex(h => h.includes("lat"));
      const lngC = hdrs.findIndex(h => h.includes("lon") || h.includes("lng"));
      if (aC < 0 || latC < 0 || lngC < 0) { alert("Could not find Address, Latitude, Longitude columns"); return; }
      const newDb = { ...geoDb };
      let count = 0;
      rows.slice(1).forEach(r => {
        const addr = String(r[aC] || "").trim();
        const lat  = parseFloat(r[latC]);
        const lng  = parseFloat(r[lngC]);
        if (addr && !isNaN(lat) && !isNaN(lng)) {
          newDb[addrKey(addr)] = { lat, lng };
          count++;
        }
      });
      saveGeoDb(newDb);
      // Patch any loaded stops directly with their new lat/lng
      setStops(prev => prev.map(s => {
        const coords = newDb[addrKey(s.address)];
        return coords ? { ...s, geocoded: { ...coords, fromDb: true } } : s;
      }));
      alert(`✓ Imported ${count} coordinates into Geo DB`);
    };
    reader.readAsBinaryString(file);
  };

  const GeoDbTab = () => {
    // Merge: all stops + all DB entries
    const stopAddrs = stops.map(s => s.address);
    const dbAddrs = Object.keys(geoDb).map(k =>
      stopAddrs.find(a => addrKey(a) === k) || k
    );
    const allAddrs = [...new Set([...stopAddrs, ...dbAddrs])];

    const dbCount = Object.keys(geoDb).length;

    return (
      <div style={{ padding: "14px" }}>

        {/* Header card */}
        <div style={card({ padding: "14px 16px", marginBottom: 10, background: C.amberBg, borderColor: C.amberBorder })}>
          <div style={{ fontWeight: 700, fontSize: 15, color: C.text, marginBottom: 4 }}>📌 Address Coordinate Database</div>
          <div style={{ fontSize: 13, color: "#92400E", lineHeight: 1.5 }}>
            Set exact lat/long for each address once — saved permanently. The app uses these instead of estimates.
          </div>
          <div style={{ marginTop: 8, display: "flex", gap: 10, alignItems: "center", flexWrap: "wrap" }}>
            <span style={pill(C.greenBg, C.green)}>{dbCount} saved</span>
            <span style={pill(C.s2, C.muted)}>{allAddrs.length - dbCount} pending</span>
            {dbSaved && <span style={pill(C.greenBg, C.green)}>✓ Saved</span>}
          </div>
          {/* Bulk import from spreadsheet */}
          <label htmlFor="geo-xlsx-input" style={{
            display: "flex", alignItems: "center", gap: 8, marginTop: 10,
            background: C.white, border: `1px solid ${C.amberBorder}`, borderRadius: 8,
            padding: "10px 14px", cursor: "pointer", fontSize: 13, fontWeight: 600, color: C.text,
          }}>
            📥 Import JAWS_GeoCoordinates.xlsx
            <span style={{ fontSize: 11, color: C.muted, fontWeight: 400, marginLeft: "auto" }}>bulk load all coords at once</span>
          </label>
          <input id="geo-xlsx-input" ref={geoXlsxRef} type="file" accept=".xlsx,.xls,.csv"
            style={{ display: "none" }}
            onChange={e => { if (e.target.files[0]) importGeoSpreadsheet(e.target.files[0]); e.target.value = ""; }}
          />
          <div style={{ fontSize: 12, color: C.muted, marginTop: 8 }}>
            💡 Tip: In Google Maps, long-press any location → tap the coordinates at the bottom to copy lat/long
          </div>
        </div>

        {/* Edit panel */}
        {editingAddr && (
          <div style={card({ border: `2px solid ${C.amber}`, background: C.amberBg, marginBottom: 10 })}>
            <div style={{ fontWeight: 700, fontSize: 14, marginBottom: 4 }}>Pinning:</div>
            <div style={{ fontFamily: FONTS.mono, fontSize: 12, color: C.muted, marginBottom: 10, wordBreak: "break-word" }}>{editingAddr}</div>

            {/* Google Maps link — opens in browser at the pinned/address location */}
            {(() => {
              const stored = geoDb[addrKey(editingAddr)];
              const hasCoords = !isNaN(parseFloat(latInput)) && !isNaN(parseFloat(lngInput));
              const lat = parseFloat(latInput) || stored?.lat;
              const lng = parseFloat(lngInput) || stored?.lng;
              const mapsUrl = (hasCoords || stored)
                ? `https://maps.google.com/?q=${lat},${lng}`
                : `https://maps.google.com/?q=${encodeURIComponent(editingAddr)}`;
              return (
                <div style={{ background: C.white, border: `1px solid ${C.border}`, borderRadius: 8, padding: "12px 14px", marginBottom: 12 }}>
                  <div style={{ fontSize: 12, color: C.muted, marginBottom: 10, lineHeight: 1.5 }}>
                    Open Google Maps to find the exact pickup location, long-press the spot, then copy the coordinates that appear at the bottom.
                  </div>
                  <a href={mapsUrl} target="_blank" rel="noopener noreferrer"
                    style={{ ...btn(C.blue, { width: "100%", color: "#fff", textDecoration: "none", fontSize: 14, padding: "11px 16px" }) }}>
                    🗺️ Open {(hasCoords || stored) ? `${lat?.toFixed(4)}, ${lng?.toFixed(4)}` : editingAddr.split(",")[0]} in Google Maps
                  </a>
                  {(hasCoords || stored) && (
                    <div style={{ marginTop: 8, fontFamily: FONTS.mono, fontSize: 11, color: C.green, textAlign: "center" }}>
                      ✓ Coordinates set · map will open at this exact pin
                    </div>
                  )}
                </div>
              );
            })()}

            {/* Step-by-step instructions */}
            <div style={{ background: C.white, border: `1px solid ${C.border}`, borderRadius: 8, padding: "10px 12px", marginBottom: 12 }}>
              <div style={{ fontWeight: 700, fontSize: 12, color: C.text, marginBottom: 6 }}>How to get exact coordinates:</div>
              {[
                ["1", "Tap the 🗺️ button next to the address in the Route tab to open Google Maps"],
                ["2", "Long-press the exact pickup spot (driveway, bin area, gate)"],
                ["3", "Tap the lat/long that appears at the bottom of the screen to copy it"],
                ["4", "Paste into the fields below — format: 30.385432, -86.384521"],
              ].map(([n, text]) => (
                <div key={n} style={{ display: "flex", gap: 8, marginBottom: 5, fontSize: 12, color: C.muted, lineHeight: 1.4 }}>
                  <span style={{ background: C.amber, color: "#000", borderRadius: "50%", width: 17, height: 17, display: "flex", alignItems: "center", justifyContent: "center", fontWeight: 700, fontSize: 10, flexShrink: 0, marginTop: 1 }}>{n}</span>
                  {text}
                </div>
              ))}
            </div>

            <div style={{ display: "flex", gap: 8, marginBottom: 6 }}>
              <div style={{ flex: 1 }}>
                <span style={lbl}>Latitude</span>
                <input value={latInput} onChange={e => setLatInput(e.target.value)}
                  onPaste={e => {
                    const text = e.clipboardData.getData("text").trim();
                    const match = text.match(/^(-?\d+\.\d+)[,\s]+(-?\d+\.\d+)$/);
                    if (match) {
                      e.preventDefault();
                      setLatInput(match[1]);
                      setLngInput(match[2]);
                    }
                  }}
                  placeholder="30.385432"
                  style={{ width: "100%", boxSizing: "border-box", background: C.white, border: `1px solid ${C.border}`, borderRadius: 7, padding: "8px 10px", fontSize: 14, fontFamily: FONTS.mono, outline: "none" }} />
              </div>
              <div style={{ flex: 1 }}>
                <span style={lbl}>Longitude</span>
                <input value={lngInput} onChange={e => setLngInput(e.target.value)}
                  placeholder="-86.384521"
                  style={{ width: "100%", boxSizing: "border-box", background: C.white, border: `1px solid ${C.border}`, borderRadius: 7, padding: "8px 10px", fontSize: 14, fontFamily: FONTS.mono, outline: "none" }} />
              </div>
            </div>
            <div style={{ fontSize: 11, color: C.muted, marginBottom: 10 }}>
              💡 You can paste "30.385432, -86.384521" directly into the Latitude field — it'll split automatically
            </div>

            <div style={{ display: "flex", gap: 8 }}>
              <button onClick={saveCoord}
                disabled={!latInput || !lngInput}
                style={{ ...btn(C.green, { flex: 2, color: "#fff", opacity: (!latInput || !lngInput) ? 0.5 : 1 }) }}>
                💾 Save coordinates
              </button>
              <button onClick={() => { setEditingAddr(null); setLatInput(""); setLngInput(""); }}
                style={{ ...btn(C.s2, { flex: 1, border: `1px solid ${C.border}`, color: C.muted }) }}>
                Cancel
              </button>
            </div>
          </div>
        )}

        {!allAddrs.length && (
          <div style={{ textAlign: "center", color: C.muted, padding: "40px 20px" }}>
            <div style={{ fontSize: 32 }}>📋</div>
            <div style={{ fontWeight: 600, marginTop: 10 }}>Load a route in Setup first</div>
            <div style={{ fontSize: 13, marginTop: 4 }}>All addresses from your spreadsheet will appear here</div>
          </div>
        )}

        {allAddrs.length > 0 && (
          <>
            <span style={lbl}>{allAddrs.length} addresses</span>
            {allAddrs.map(addr => {
              const key = addrKey(addr);
              const stored = geoDb[key];
              const isEditing = editingAddr === addr;
              return (
                <div key={key} style={card({
                  padding: "10px 12px", marginBottom: 7,
                  borderLeft: `4px solid ${stored ? C.green : C.border}`,
                  opacity: isEditing ? 1 : 1,
                })}>
                  <div style={{ display: "flex", alignItems: "flex-start", gap: 10 }}>
                    <div style={{ flex: 1, minWidth: 0 }}>
                      <div style={{ fontSize: 13, fontWeight: 600, marginBottom: 3, wordBreak: "break-word" }}>{addr}</div>
                      {stored ? (
                        <div style={{ fontFamily: FONTS.mono, fontSize: 11, color: C.green }}>
                          ✓ {stored.lat.toFixed(5)}, {stored.lng.toFixed(5)}
                        </div>
                      ) : (
                        <div style={{ fontSize: 11, color: C.amber }}>
                          No match — key: <span style={{ fontFamily: FONTS.mono }}>{addrKey(addr).slice(0, 50)}</span>
                        </div>
                      )}
                    </div>
                    <div style={{ display: "flex", gap: 6, flexShrink: 0 }}>
                      <button onClick={() => {
                        setEditingAddr(addr);
                        setLatInput(stored ? stored.lat.toFixed(6) : "");
                        setLngInput(stored ? stored.lng.toFixed(6) : "");
                      }} style={{ ...btn(stored ? C.s2 : C.amber, { padding: "5px 10px", fontSize: 12, border: `1px solid ${stored ? C.border : C.amber}`, color: stored ? C.text : "#000" }) }}>
                        {stored ? "✎ Edit" : "📌 Pin"}
                      </button>
                      {stored && (
                        <button onClick={() => deleteCoord(addr)}
                          style={{ ...btn(C.s2, { padding: "5px 8px", fontSize: 12, border: `1px solid ${C.border}`, color: C.red }) }}>
                          ✕
                        </button>
                      )}
                    </div>
                  </div>
                </div>
              );
            })}
          </>
        )}
      </div>
    );
  };

  // ─── ROOT RENDER ────────────────────────────────────────────────────────────

  return (
    <div style={{ background: C.bg, minHeight: "100vh", color: C.text, fontFamily: FONTS.sans, maxWidth: 600, margin: "0 auto", paddingBottom: 40 }}>
      <input ref={photoRef} type="file" accept="image/*" style={{ display: "none" }} onChange={handlePhoto} />
      <input id="xlsx-input" ref={xlsxRef} type="file" accept=".xlsx,.xls,.csv" style={{ display: "none" }} onChange={e => e.target.files[0] && parseExcel(e.target.files[0])} />

      {/* HEADER */}
      <div style={{
        background: C.text, padding: "14px 18px", display: "flex",
        alignItems: "center", justifyContent: "space-between",
        position: "sticky", top: 0, zIndex: 100, borderBottom: `3px solid ${C.amber}`,
      }}>
        <div style={{ display: "flex", alignItems: "center", gap: 12 }}>
          <div style={{
            width: 40, height: 40, borderRadius: "50%", background: C.amber,
            display: "flex", alignItems: "center", justifyContent: "center", fontSize: 22,
          }}>🦈</div>
          <div>
            <div style={{ fontSize: 22, fontWeight: 800, color: C.white, letterSpacing: 2, lineHeight: 1 }}>JAWS</div>
            <div style={{ fontSize: 10, color: "#9CA3AF", letterSpacing: 3, fontWeight: 600 }}>SERVICES INC.</div>
          </div>
        </div>
        {stops.length > 0 && (
          <div style={{ textAlign: "right" }}>
            <div style={{ fontFamily: FONTS.mono, fontSize: 22, fontWeight: 800, color: allDone ? C.green : C.white, lineHeight: 1 }}>
              {completedCount}/{stops.length}
            </div>
            <div style={{ fontSize: 10, color: "#9CA3AF", letterSpacing: 2 }}>COLLECTED</div>
          </div>
        )}
      </div>

      {/* Shared route banner — shown when a route has been published */}
      {sharedRoute && !stops.length && (
        <div style={{
          background: "#EFF6FF", borderBottom: "1px solid #BFDBFE",
          padding: "10px 16px", display: "flex", alignItems: "center", justifyContent: "space-between", gap: 10,
        }}>
          <div>
            <div style={{ fontSize: 13, fontWeight: 700, color: "#1E40AF" }}>
              📡 Route ready — {sharedRoute.count} stops
            </div>
            <div style={{ fontSize: 11, color: "#3B82F6" }}>Published {sharedRoute.publishedAt}</div>
          </div>
          <button
            onClick={loadSharedRoute}
            disabled={checkingShared}
            style={{ ...btn("#3B82F6", { color: "#fff", fontSize: 13, padding: "8px 14px", flexShrink: 0 }) }}
          >
            {checkingShared ? "Loading..." : "Load route →"}
          </button>
        </div>
      )}
      <div style={{ display: "flex", background: C.white, borderBottom: `1px solid ${C.border}`, position: "sticky", top: 67, zIndex: 99 }}>
        {[
          { id: "setup",   icon: "📋", label: "Setup" },
          { id: "route",   icon: "🗺️", label: "Route" },
          { id: "driver",  icon: "🚛", label: "Driver" },
          { id: "geodb",   icon: "📌", label: "Geo DB" },
          { id: "summary", icon: "📊", label: "Summary" },
        ].map(t => (
          <button key={t.id} onClick={() => setTab(t.id)} style={{
            flex: 1, padding: "10px 2px", background: "transparent", border: "none",
            borderBottom: tab === t.id ? `3px solid ${C.amber}` : "3px solid transparent",
            color: tab === t.id ? C.text : C.muted, fontSize: 11, fontWeight: tab === t.id ? 700 : 500,
            cursor: "pointer", fontFamily: FONTS.sans, marginBottom: -1,
            display: "flex", alignItems: "center", justifyContent: "center", gap: 5,
          }}>
            <span>{t.icon}</span> {t.label}
          </button>
        ))}
      </div>

      {tab === "setup"   && SetupTab()}
      {tab === "route"   && RouteTab()}
      {tab === "driver"  && DriverTab()}
      {tab === "geodb"   && GeoDbTab()}
      {tab === "summary" && SummaryTab()}

      <style>{`@keyframes gps-pulse { 0%,100%{opacity:1} 50%{opacity:0.4} }`}</style>
    </div>
  );
}
