const API_BASE_URL = import.meta.env.VITE_API_BASE_URL;

export async function uploadFile(file) {
  const formData = new FormData();
  formData.append("file", file);

  const response = await fetch(`${API_BASE_URL}/upload`, {
    method: "POST",
    body: formData,
  });

  if (!response.ok) {
    const error = await response.json();
    throw new Error(error.detail || "Upload failed");
  }

  return response.json();
}

export async function processFile(fileId) {
  const response = await fetch(`${API_BASE_URL}/process/${fileId}`, {
    method: "POST",
  });

  if (!response.ok) {
    const error = await response.json();
    throw new Error(error.detail || "Processing failed");
  }

  return response.json();
}

export function getFilledDownloadUrl(fileId) {
  return `${API_BASE_URL}/download/filled/${fileId}`;
}

export function getReviewLogDownloadUrl(fileId) {
  return `${API_BASE_URL}/download/log/${fileId}`;
}