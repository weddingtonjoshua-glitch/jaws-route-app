// app/api/route/route.js
// Upstash Redis-backed shared route with real-time stop status sync

const REDIS_URL   = process.env.UPSTASH_REDIS_REST_URL;
const REDIS_TOKEN = process.env.UPSTASH_REDIS_REST_TOKEN;
const ROUTE_KEY   = "jaws:route";

async function redisCmd(...args) {
  const res = await fetch(`${REDIS_URL}/${args.map(encodeURIComponent).join("/")}`, {
    headers: { Authorization: `Bearer ${REDIS_TOKEN}` },
  });
  return res.json();
}

async function getRoute() {
  const data = await redisCmd("GET", ROUTE_KEY);
  if (!data?.result) return null;
  return JSON.parse(data.result);
}

async function setRoute(route) {
  await redisCmd("SET", ROUTE_KEY, JSON.stringify(route));
}

// GET — return current route
export async function GET() {
  try {
    const route = await getRoute();
    if (!route) return Response.json({ route: null });
    return Response.json({ route });
  } catch (e) {
    return Response.json({ error: e.message }, { status: 500 });
  }
}

// POST — publish a new route (dispatch)
export async function POST(req) {
  try {
    const body = await req.json();
    await setRoute(body);
    return Response.json({ ok: true });
  } catch (e) {
    return Response.json({ error: e.message }, { status: 500 });
  }
}

// PATCH — update a single stop's status (driver)
export async function PATCH(req) {
  try {
    const { id, patch } = await req.json(); // patch = { status, timestamp } or { status: "claimed" }
    const route = await getRoute();
    if (!route) return Response.json({ error: "No route" }, { status: 404 });
    route.stops = route.stops.map(s => s.id === id ? { ...s, ...patch } : s);
    await setRoute(route);
    return Response.json({ ok: true });
  } catch (e) {
    return Response.json({ error: e.message }, { status: 500 });
  }
}

// DELETE — clear the route (after season / end of day)
export async function DELETE() {
  try {
    await redisCmd("DEL", ROUTE_KEY);
    return Response.json({ ok: true });
  } catch (e) {
    return Response.json({ error: e.message }, { status: 500 });
  }
}
