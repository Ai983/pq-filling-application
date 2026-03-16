import { useState } from "react";

export default function FileUpload({ onSubmit, loading }) {
  const [selectedFile, setSelectedFile] = useState(null);

  function handleFileChange(event) {
    const file = event.target.files?.[0];
    if (file) {
      setSelectedFile(file);
    }
  }

  function handleSubmit(event) {
    event.preventDefault();

    if (!selectedFile) {
      alert("Please select an Excel file first.");
      return;
    }

    onSubmit(selectedFile);
  }

  return (
    <form onSubmit={handleSubmit} className="card">
      <h2>Upload Vendor Registration File</h2>
      <p>Select an Excel file and process it through the autofill backend.</p>

      <input
        type="file"
        accept=".xlsx"
        onChange={handleFileChange}
        disabled={loading}
      />

      {selectedFile && (
        <div className="file-info">
          <strong>Selected:</strong> {selectedFile.name}
        </div>
      )}

      <button type="submit" disabled={loading}>
        {loading ? "Processing..." : "Upload & Process"}
      </button>
    </form>
  );
}