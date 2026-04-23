//! Diff engine — pure logic, no I/O.

pub mod token;
pub mod diff;
pub mod stats;
pub mod pair;
pub mod block;
pub mod block_diff;

pub use diff::{diff_words, diff_paragraphs, Change, ChangeKind, ParaOp};
pub use stats::{Stats, stats_of, stats_of_blocks};
pub use pair::{detect_pairs, PairDetection, PairMatch, Marker, StateTag};
pub use block::{Block, RichParagraph, RichRun, TableCell, TableRow};
pub use block_diff::{diff_blocks, BlockOp, CellOp, RowOp, TableDiff};
