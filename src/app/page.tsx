import { getServerSession } from "next-auth";
import { authOptions } from "@/lib/auth";
import { getHorses, scanState, unlinkedFolders } from "@/lib/sharepoint";
import { SEASON } from "@/lib/season";
import App from "./App";
import SignIn from "./SignIn";
import Providers from "./Providers";

// Always current. The data is small and Philip opens this a few times a day.
export const dynamic = "force-dynamic";

export default async function Page() {
  const session = await getServerSession(authOptions);

  if (!session || (session as any).error) {
    return (
      <Providers>
        <SignIn
          note={
            (session as any)?.error
              ? "Your session expired. Sign in again."
              : undefined
          }
        />
      </Providers>
    );
  }

  try {
    const horses = await getHorses();
    const folders = await unlinkedFolders(horses);
    return (
      <Providers>
        <App
          horses={horses}
          season={SEASON}
          rescan={scanState().stale}
          folders={folders}
        />
      </Providers>
    );
  } catch (e: any) {
    if (e?.message === "NOT_SIGNED_IN") {
      return (
        <Providers>
          <SignIn note="Your session expired. Sign in again." />
        </Providers>
      );
    }
    return (
      <Providers>
        <SignIn note="Couldn't reach SharePoint just now. Try again in a moment." />
      </Providers>
    );
  }
}
