import { useState } from 'react';
import './App.scss';

function App() {
  const [publicHolidayFile, setPublicHolidayFile] = useState<File | null>(null);
  const [vacationDaysFile, setVacationDaysFile] = useState<File | null>(null);
  const [timeDistributionFile, setTimeDistributionFile] = useState<File | null>(null);

  const handleFileChange = (setter: React.Dispatch<React.SetStateAction<File | null>>) =>
    (e: React.ChangeEvent<HTMLInputElement>) => {
      if (e.target.files) setter(e.target.files[0]);
    };

  const handleUploadAndProcessAll = () => {
    // Implement upload & processing logic
    console.log('Processing all files:', {
      publicHolidayFile,
      vacationDaysFile,
      timeDistributionFile,
    });
  };

  return (
    <div className="container">
      <h1>Time Management File Processor</h1>

      <section>
        <h2>Public Holidays</h2>
        <p>Upload the Excel file with public holiday data.</p>
        <input type="file" accept=".xlsx" onChange={handleFileChange(setPublicHolidayFile)} />
        {publicHolidayFile && <p>{publicHolidayFile.name}</p>}
        <a href="#" className="download-link">Download Processed File</a>
      </section>

      <section>
        <h2>Vacation Days</h2>
        <p>Upload vacation days for employees</p>
        <input type="file" accept=".xlsx" onChange={handleFileChange(setVacationDaysFile)} />
        {vacationDaysFile && <p>{vacationDaysFile.name}</p>}
        <a href="#" className="download-link">Download Processed File</a>
      </section>

      <section>
        <h2>Time Distribution</h2>
        <p>Upload timesheet or work hours distribution</p>
        <input type="file" accept=".xlsx" onChange={handleFileChange(setTimeDistributionFile)} />
        {timeDistributionFile && <p>{timeDistributionFile.name}</p>}
        <a href="#" className="download-link">Download Processed File</a>
      </section>

      <button className="process-button" onClick={handleUploadAndProcessAll}>
        Upload & Process All
      </button>
    </div>
  );
}

export default App;