// app/api/route/route.js
// Upstash Redis-backed shared route with real-time stop status sync

const REDIS_URL   = process.env.UPSTASH_REDIS_REST_URL;
const REDIS_TOKEN = process.env.UPSTASH_REDIS_REST_TOKEN;
const ROUTE_KEY   = "jaws:route";

// Use POST pipeline for all commands — handles large payloads safely
async function redisPost(command) {
  const res = await fetch(REDIS_URL, {
    method: "POST",
    headers: {
      Authorization: `Bearer ${REDIS_TOKEN}`,
      "Content-Type": "application/json",
    },
    body: JSON.stringify(command),
  });
  return res.json();
}

async function getRoute() {
  const data = await redisPost(["GET", ROUTE_KEY]);
  if (!data?.result) return null;
  return JSON.parse(data.result);
}

async function setRoute(route) {
  await redisPost(["SET", ROUTE_KEY, JSON.stringify(route)]);
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
    const { id, patch } = await req.json();
    const route = await getRoute();
    if (!route) return Response.json({ error: "No route" }, { status: 404 });
    route.stops = route.stops.map(s => s.id === id ? { ...s, ...patch } : s);
    await setRoute(route);
    return Response.json({ ok: true });
  } catch (e) {
    return Response.json({ error: e.message }, { status: 500 });
  }
}

// DELETE — clear the route
export async function DELETE() {
  try {
    await redisPost(["DEL", ROUTE_KEY]);
    return Response.json({ ok: true });
  } catch (e) {
    return Response.json({ error: e.message }, { status: 500 });
  }
}
