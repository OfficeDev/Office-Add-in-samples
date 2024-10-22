export class FallbackError extends Error {
  constructor(message: string) {
    super(message);
    this.name = "FallbackError";
  }
}
