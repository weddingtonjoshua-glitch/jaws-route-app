let sharedRoute = null;

export async function GET() {
  return Response.json({ route: sharedRoute });
}

export async function POST(request) {
  try {
    const body = await request.json();
    sharedRoute = body;
    return Response.json({ ok: true });
  } catch {
    return Response.json({ error: "Invalid body" }, { status: 400 });
  }
}
