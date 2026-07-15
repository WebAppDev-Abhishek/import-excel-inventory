import { ExcelDropzone } from '@/components/ExcelDropzone';

export default function Home() {
  return (
    <main className="page-shell">
      <section className="hero">
        <p className="eyebrow">Inventory preview</p>
        <h1>Drop an Excel file and preview every inventory row as a card.</h1>
        <p className="hero-copy">
          This demo parses the workbook in the browser, keeps the UI responsive, and adapts to whatever
          columns appear in the sheet.
        </p>
      </section>

      <ExcelDropzone />
    </main>
  );
}
