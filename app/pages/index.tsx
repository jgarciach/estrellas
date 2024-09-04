import CSVToDocxConverter from '../components/CSVToDocxConverter';

export default function Home() {
  return (
    <div className="container mx-auto px-4">
      <h1 className="text-2xl font-bold my-4">CSV to DOCX Converter</h1>
      <CSVToDocxConverter />
    </div>
  );
}