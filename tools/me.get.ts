import { graphGet } from "../lib/graph.ts";

interface GraphUser {
  id: string;
  displayName?: string;
  givenName?: string;
  surname?: string;
  userPrincipalName?: string;
  mail?: string;
  jobTitle?: string;
  mobilePhone?: string;
  officeLocation?: string;
  preferredLanguage?: string;
}

/**
 * @effect readOnly
 * @idempotent
 */
export default async function tool(): Promise<unknown> {
  return await graphGet<GraphUser>("/me");
}
