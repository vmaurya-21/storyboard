// Mock @microsoft/sp-core-library
export class Version {
  public major: number = 1;
  public minor: number = 0;
  public patch: number = 0;

  constructor(major: number = 1, minor: number = 0, patch: number = 0) {
    this.major = major;
    this.minor = minor;
    this.patch = patch;
  }

  public toString(): string {
    return `${this.major}.${this.minor}`;
  }

  public static parse(version: string): Version {
    const parts = version.split('.');
    return new Version(
      parseInt(parts[0] || '1', 10),
      parseInt(parts[1] || '0', 10),
      parseInt(parts[2] || '0', 10)
    );
  }
}

export const Environment = {
  type: 1
};

export const EnvironmentType = {
  Local: 0,
  SharePoint: 1,
  ClassicSharePoint: 2
};

export const Log = {
  info: jest.fn(),
  warn: jest.fn(),
  error: jest.fn(),
  verbose: jest.fn()
};
