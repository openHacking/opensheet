type Fingerprint = { algorithm: string; files: number; sha256: string };
export function sourceFingerprint(): Fingerprint;
export function localBuildFingerprint(directory?: string): Fingerprint;
