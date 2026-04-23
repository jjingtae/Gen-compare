//! DOCX read/write. MVP: extract paragraph text + write redline DOCX.

pub mod read;
pub mod redline;

pub use read::{read_paragraphs, read_document, DocxContent, Paragraph, RawParts};
pub use redline::{write_redline, RedlineOptions, RedlineStyle};
