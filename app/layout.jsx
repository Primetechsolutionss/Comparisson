export const metadata = {
  title: 'Leveransplan Validator',
  description: 'Document delivery control — compare delivery sheets against master Leveransplan',
};

export default function RootLayout({ children }) {
  return (
    <html lang="sv">
      <body style={{ margin: 0, padding: 0, background: '#0B1120' }}>
        {children}
      </body>
    </html>
  );
}
