import type { Metadata } from "next";
import "./globals.css";

const appTitle = process.env.NEXT_PUBLIC_APP_TITLE ?? "CAFE SHIFT";

export const metadata: Metadata = {
  title: appTitle,
  description: "CAFE SHIFT parent app for the GAS shift management web app",
};

export default function RootLayout({
  children,
}: Readonly<{
  children: React.ReactNode;
}>) {
  return (
    <html lang="ja">
      <body>{children}</body>
    </html>
  );
}
