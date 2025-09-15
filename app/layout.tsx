// app/layout.tsx
import './globals.css'; // <- IMPORTA O CSS GLOBAL

export const metadata = {
  title: 'Conferência de Inventário — Físico x SAP',
  description: 'App de conferência Físico x SAP',
};

export default function RootLayout({
  children,
}: {
  children: React.ReactNode
}) {
  // data-theme="cards" = tema padrão (Mockup Cards). O botão troca para "dark".
  return (
    <html lang="pt-BR" data-theme="cards">
      <body>{children}</body>
    </html>
  );
}
