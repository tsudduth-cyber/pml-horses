import type { Metadata, Viewport } from "next";
import "./globals.css";

export const metadata: Metadata = {
  title: "PML Performance Horses",
  description: "Philip's horse records, on his phone.",
  manifest: "/manifest.json",
  appleWebApp: {
    capable: true,
    title: "PML",
    statusBarStyle: "default",
  },
  icons: {
    apple: "/icon-180.png",
    icon: "/icon-192.png",
  },
};

export const viewport: Viewport = {
  themeColor: "#EDE7DE",
  width: "device-width",
  initialScale: 1,
};

export default function RootLayout({
  children,
}: {
  children: React.ReactNode;
}) {
  return (
    <html lang="en">
      <head>
        <link rel="preconnect" href="https://fonts.googleapis.com" />
        <link
          rel="preconnect"
          href="https://fonts.gstatic.com"
          crossOrigin=""
        />
        <link
          href="https://fonts.googleapis.com/css2?family=Cormorant+Garamond:wght@500;600;700&family=Karla:wght@400;500;600;700&display=swap"
          rel="stylesheet"
        />
      </head>
      <body>{children}</body>
    </html>
  );
}
