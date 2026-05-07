import './styles.css';

export const metadata = {
  title: 'docx-sax Next.js WASM demo',
  description: 'Upload a DOCX and preview text parsed by docx-sax/browser in WebAssembly.',
};

export default function RootLayout({ children }) {
  return (
    <html lang="en">
      <body>{children}</body>
    </html>
  );
}
