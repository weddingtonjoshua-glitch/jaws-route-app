export const metadata = { title: "JAWS Services", description: "Route management" };
export default function RootLayout({ children }) {
  return (
    <html lang="en">
      <head>
        <meta name="viewport" content="width=device-width, initial-scale=1, maximum-scale=1" />
        <meta name="mobile-web-app-capable" content="yes" />
        <meta name="apple-mobile-web-app-capable" content="yes" />
        <meta name="apple-mobile-web-app-title" content="JAWS Routes" />
      </head>
      <body style={{ margin: 0, padding: 0, background: "#F8F9FA" }}>{children}</body>
    </html>
  );
}
