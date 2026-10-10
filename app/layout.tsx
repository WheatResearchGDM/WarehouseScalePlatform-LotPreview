import type { Metadata } from "next";
import "./globals.css";

export const metadata: Metadata = {
  title: "Trial Weighing | GDM",
  description:
    "Plot scanning, plot or seed weight recording, and wheat trial weighing progress.",
  icons: {
    icon: "/favicon.svg",
    shortcut: "/favicon.svg",
  },
};

export default function RootLayout({
  children,
}: Readonly<{
  children: React.ReactNode;
}>) {
  return (
    <html lang="en-US">
      <body className="antialiased">{children}</body>
    </html>
  );
}
