export class CalculationViewRetired extends Error {
  constructor(cause: unknown) {
    super('The calculation view changed.', { cause });
  }
}
