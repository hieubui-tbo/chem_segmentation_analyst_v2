import './globals.css';

export const metadata = {
  title: 'ChemSeg Analyst — Chemical Segmentation Tool',
  description: 'V12 Chemical Import/Export Segmentation & Price Analysis Tool',
};

export default function RootLayout({ children }) {
  return (
    <html lang="en">
      <body>{children}</body>
    </html>
  );
}
