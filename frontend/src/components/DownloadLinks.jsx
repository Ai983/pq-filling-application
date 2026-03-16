import {
  getFilledDownloadUrl,
  getReviewLogDownloadUrl,
} from "../api/client";

export default function DownloadLinks({ fileId }) {
  if (!fileId) return null;

  return (
    <div className="card">
      <h2>Downloads</h2>
      <div className="download-actions">
        <a
          href={getFilledDownloadUrl(fileId)}
          target="_blank"
          rel="noreferrer"
          className="download-button"
        >
          Download Filled File
        </a>

        <a
          href={getReviewLogDownloadUrl(fileId)}
          target="_blank"
          rel="noreferrer"
          className="download-button secondary"
        >
          Download Review Log
        </a>
      </div>
    </div>
  );
}