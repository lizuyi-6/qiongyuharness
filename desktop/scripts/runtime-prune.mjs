const DEV_ONLY_SUFFIXES = [".ts", ".mts", ".cts", ".map"];

export function isDevOnlyRuntimeFile(fileName) {
  return typeof fileName === "string" && DEV_ONLY_SUFFIXES.some((suffix) => fileName.endsWith(suffix));
}
