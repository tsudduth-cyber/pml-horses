import type { NextAuthOptions } from "next-auth";
import AzureADProvider from "next-auth/providers/azure-ad";

/**
 * Delegated Graph access. Philip signs in with his own Microsoft account and
 * the app sees exactly what he can already see in SharePoint — nothing more.
 */
/**
 * Day-to-day the app only reads horse folders and writes care dates, so
 * Sites.ReadWrite.All is all it asks for.
 *
 * Creating the "PML Horses" list needed Sites.Manage.All, because
 * Sites.ReadWrite.All only covers items inside a list that already exists.
 * That was a one-time job and the scope is deliberately not requested here
 * any more. If the list ever has to be rebuilt, revert this commit to put
 * "Sites.Manage.All" back, run /api/setup?confirm=yes, then revert again.
 */
const SCOPES = [
  "openid",
  "profile",
  "email",
  "offline_access",
  "Sites.ReadWrite.All",
].join(" ");

const TENANT = process.env.AZURE_AD_TENANT_ID!;

async function refresh(token: any) {
  try {
    const res = await fetch(
      `https://login.microsoftonline.com/${TENANT}/oauth2/v2.0/token`,
      {
        method: "POST",
        headers: { "Content-Type": "application/x-www-form-urlencoded" },
        body: new URLSearchParams({
          client_id: process.env.AZURE_AD_CLIENT_ID!,
          client_secret: process.env.AZURE_AD_CLIENT_SECRET!,
          grant_type: "refresh_token",
          refresh_token: token.refreshToken,
          scope: SCOPES,
        }),
      }
    );
    const data = await res.json();
    if (!res.ok) throw data;
    return {
      ...token,
      accessToken: data.access_token,
      expiresAt: Date.now() + (data.expires_in ?? 3600) * 1000,
      // Entra does not always return a new refresh token; keep the old one.
      refreshToken: data.refresh_token ?? token.refreshToken,
      error: undefined,
    };
  } catch {
    // Surface it rather than silently serving a dead token.
    return { ...token, error: "RefreshFailed" };
  }
}

export const authOptions: NextAuthOptions = {
  providers: [
    AzureADProvider({
      clientId: process.env.AZURE_AD_CLIENT_ID!,
      clientSecret: process.env.AZURE_AD_CLIENT_SECRET!,
      tenantId: TENANT,
      authorization: { params: { scope: SCOPES } },
    }),
  ],
  session: { strategy: "jwt", maxAge: 60 * 60 * 24 * 30 },
  callbacks: {
    async jwt({ token, account }) {
      if (account) {
        return {
          ...token,
          accessToken: account.access_token,
          refreshToken: account.refresh_token,
          expiresAt: (account.expires_at ?? 0) * 1000,
        };
      }
      // Renew a minute before it actually expires.
      if (Date.now() < (token as any).expiresAt - 60_000) return token;
      return refresh(token);
    },
    async session({ session, token }) {
      (session as any).error = (token as any).error;
      return session;
    },
  },
};

/** Server-side only. Never reaches the browser. */
export async function getAccessToken(): Promise<string> {
  const { getServerSession } = await import("next-auth");
  const { getToken } = await import("next-auth/jwt");
  const { cookies } = await import("next/headers");

  const session = await getServerSession(authOptions);
  if (!session) throw new Error("NOT_SIGNED_IN");

  const jar = await cookies();
  const raw = Object.fromEntries(jar.getAll().map((c) => [c.name, c.value]));
  const token: any = await getToken({
    req: { cookies: raw, headers: new Headers() } as any,
    secret: process.env.NEXTAUTH_SECRET!,
  });

  if (!token?.accessToken || token.error) throw new Error("NOT_SIGNED_IN");
  return token.accessToken as string;
}
