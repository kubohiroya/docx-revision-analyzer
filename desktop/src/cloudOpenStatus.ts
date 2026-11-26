import type { CloudOpenStage } from "./shared";

export class MicrosoftError extends Error {
  constructor(
    message: string,
    /** 画面で案内を出し分けるための種類 */
    readonly code:
      | "notConfigured"
      | "badUrl"
      | "notDocx"
      | "notFolder"
      | "folder"
      | "notFound"
      | "forbidden"
      | "tooLarge"
      | "signInCancelled"
      | "network",
    readonly stage?: CloudOpenStage
  ) {
    super(message);
  }
}

/** Keep the useful classification and stage without exposing service URLs or tokens. */
export async function atCloudOpenStage<T>(stage: CloudOpenStage, run: () => Promise<T>): Promise<T> {
  try {
    return await run();
  } catch (error) {
    throw new MicrosoftError("", error instanceof MicrosoftError ? error.code : "network", stage);
  }
}
