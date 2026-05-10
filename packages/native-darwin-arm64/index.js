export const platform = 'darwin-arm64';
export const implemented = false;

export function unavailable() {
  throw new Error('@docx-sax/native-darwin-arm64 is a placeholder prerelease package for trusted-publisher/bootstrap only. It does not include a native runtime implementation yet.');
}

export default {
  platform,
  implemented,
  unavailable,
};
