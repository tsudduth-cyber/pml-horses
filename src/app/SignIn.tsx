"use client";

import { signIn } from "next-auth/react";

export default function SignIn({ note }: { note?: string }) {
  return (
    <div className="gate">
      {/* eslint-disable-next-line @next/next/no-img-element */}
      <img src="/pml-logo.png" alt="PML Performance Horses" />
      <p>{note ?? "Sign in with your Microsoft account to see the horses."}</p>
      <button className="btn" onClick={() => signIn("azure-ad")}>
        Sign in
      </button>
    </div>
  );
}
