export const metadata = { title: "PROF INSTAL", description: "Asystent CWU" };

export default function RootLayout({ children }: { children: React.ReactNode }) {
  return (
    <html lang="pl">
      <body style={{ margin: 0, fontFamily: "system-ui, Arial" }}>{children}</body>
    </html>
  );
}