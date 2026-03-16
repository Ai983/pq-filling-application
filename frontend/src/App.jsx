import { useState } from "react";
import FileUpload from "./components/FileUpload";
import StatusCard from "./components/StatusCard";
import DownloadLinks from "./components/DownloadLinks";
import { uploadFile, processFile } from "./api/client";

export default function App() {
  const [loading, setLoading] = useState(false);
  const [uploadData, setUploadData] = useState(null);
  const [processData, setProcessData] = useState(null);
  const [error, setError] = useState("");

  async function handleUploadAndProcess(file) {
    try {
      setLoading(true);
      setError("");
      setUploadData(null);
      setProcessData(null);

      const uploadResponse = await uploadFile(file);
      const uploaded = uploadResponse.data;
      setUploadData(uploaded);

      const processResponse = await processFile(uploaded.file_id);
      const processed = processResponse.data;
      setProcessData(processed);
    } catch (err) {
      setError(err.message || "Something went wrong");
    } finally {
      setLoading(false);
    }
  }

  const fileId = processData?.file_id || uploadData?.file_id || null;

  return (
    <div className="app-container">
      <div className="page">
        <header className="hero">
          <h1>Vendor Autofill System</h1>
          <p>
            Upload a vendor registration Excel file, process it through the
            rule-based backend, and download the filled output and review log.
          </p>
        </header>

        <FileUpload onSubmit={handleUploadAndProcess} loading={loading} />

        <StatusCard
          uploadData={uploadData}
          processData={processData}
          error={error}
        />

        <DownloadLinks fileId={fileId} />
      </div>
    </div>
  );
}