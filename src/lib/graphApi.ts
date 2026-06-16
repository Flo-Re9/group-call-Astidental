import { graphConfig } from "./msalConfig";

export interface GroupMember {
  id: string;
  displayName: string;
  mail: string;
  jobTitle?: string;
  userPrincipalName: string;
}

export async function fetchGroupMembers(accessToken: string, endpoint: string): Promise<GroupMember[]> {
  const response = await fetch(endpoint, {
    headers: { Authorization: `Bearer ${accessToken}` },
  });

  if (!response.ok) {
    throw new Error(`Graph API error: ${response.status}`);
  }

  const data = await response.json();
  return (data.value as GroupMember[]).filter(
    (m) => m["@odata.type"] === "#microsoft.graph.user"
  );
}

export async function fetchMe(accessToken: string) {
  const response = await fetch(graphConfig.graphMeEndpoint, {
    headers: { Authorization: `Bearer ${accessToken}` },
  });
  if (!response.ok) throw new Error(`Graph API error: ${response.status}`);
  return response.json();
}

export function getTeamsCallLink(users: string, topic?: string): string {
  const base = `https://teams.microsoft.com/l/call/0/0?users=${encodeURIComponent(users)}`;
  return topic ? `${base}&topic=${encodeURIComponent(topic)}` : base;
}

export function getTeamsChatLink(emails: string[]): string {
  return `https://teams.microsoft.com/l/chat/0/0?users=${emails.map(encodeURIComponent).join(",")}`;
}
