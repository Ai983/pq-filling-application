export default function StatusCard({ uploadData, processData, error }) {
  return (
    <div className="card">
      <h2>Status</h2>

      {error && <div className="error-box">{error}</div>}

      {!error && !uploadData && !processData && (
        <p>No file processed yet.</p>
      )}

      {uploadData && (
        <div className="status-block">
          <h3>Upload Success</h3>
          <p><strong>File ID:</strong> {uploadData.file_id}</p>
          <p><strong>Original File:</strong> {uploadData.original_filename}</p>
        </div>
      )}

      {processData && (
        <div className="status-block">
          <h3>Processing Complete</h3>
          <p><strong>Total Logged Items:</strong> {processData.total_logged_items}</p>
          <p><strong>Filled:</strong> {processData.filled_count}</p>
          <p><strong>Skipped:</strong> {processData.skipped_count}</p>
          <p><strong>Unmatched:</strong> {processData.unmatched_count}</p>
        </div>
      )}
    </div>
  );
}