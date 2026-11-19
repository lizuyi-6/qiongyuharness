async function sessionFetch(session, origin, url, options = {}) {
  const cookies = await session.cookies.get({ url: origin, name: "qo_session" });
  const cookie = cookies.find((item) => item.name === "qo_session" && typeof item.value === "string" &&
    /^[A-Za-z0-9_-]{32,128}$/.test(item.value));
  if (!cookie) throw new Error("请先登录穹驭办公");

  const headers = new Headers(options.headers || {});
  headers.set("Cookie", `qo_session=${cookie.value}`);
  return session.fetch(url, { ...options, credentials: "include", headers });
}

module.exports = { sessionFetch };
