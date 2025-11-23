import I18n from "discourse-i18n";

/**
 * Get the icon name for a file based on its type/extension
 * @param {Object} file - File object from OneDrive API
 * @returns {string} Icon name for the file type
 */
export function getFileTypeIcon(file) {
  if (file.folder) {
    return "folder";
  }

  const name = file.name || "";
  const extension = name.split(".").pop()?.toLowerCase() || "";

  // Image types
  if (
    ["jpg", "jpeg", "png", "gif", "svg", "webp", "ico", "bmp"].includes(
      extension
    )
  ) {
    return "image";
  }

  // Video types
  if (["mp4", "mov", "avi", "webm", "mkv", "flv", "wmv"].includes(extension)) {
    return "video";
  }

  // Audio types
  if (["mp3", "wav", "ogg", "flac", "aac", "m4a"].includes(extension)) {
    return "music";
  }

  // Archive types
  if (["zip", "rar", "7z", "tar", "gz", "bz2", "xz"].includes(extension)) {
    return "file-zipper";
  }

  // Document types - specific icons
  if (["pdf"].includes(extension)) {
    return "file-pdf";
  }
  if (["doc", "docx"].includes(extension)) {
    return "file-word";
  }
  if (["xls", "xlsx"].includes(extension)) {
    return "file-excel";
  }
  if (["ppt", "pptx"].includes(extension)) {
    return "file-powerpoint";
  }
  if (["csv"].includes(extension)) {
    return "file-csv";
  }
  if (["txt", "md", "rtf"].includes(extension)) {
    return "file-lines";
  }

  // Default file icon
  return "file";
}

/**
 * Get the file type name/extension for display
 * @param {Object} file - File object from OneDrive API
 * @returns {string} File type name (extension or "Folder" or "File")
 */
export function getFileTypeName(file) {
  if (file.folder) {
    return "Folder";
  }

  const name = file.name || "";
  const extension = name.split(".").pop()?.toUpperCase() || "";

  if (extension) {
    return extension;
  }

  return "File";
}

/**
 * Format file size for display
 * @param {number|null|undefined} size - File size in bytes
 * @returns {string} Formatted file size or "—" if not available
 */
export function formatFileSize(size) {
  if (!size) {
    return "—";
  }
  return I18n.toHumanSize(size);
}
