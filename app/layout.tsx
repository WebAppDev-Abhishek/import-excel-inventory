import type { Metadata } from 'next';
import './globals.css';

export const metadata: Metadata = {
  title: 'Excel Inventory Preview',
  description: 'Drag and drop an Excel spreadsheet to preview inventory rows as cards.'
};

export default function RootLayout({ children }: { children: React.ReactNode }) {
  return (
    <html lang="en">
      <body>{children}</body>
    </html>
  );
}
