use std::path::PathBuf;

use clap::Parser;

/// Parse User Inputs
#[derive(Parser, Debug, Clone)]
#[command(author, version, about, long_about = None, allow_external_subcommands = true)]
pub struct Context {
    /// directory path
    #[arg(short = 'd', long = "dir", required = false, help = "top level directory to begin seach")]
    pub dir: Option<PathBuf>,
    /// remove links
    #[arg(short = 'r', long = "remove", required = false, help = "this will change the workbook structure use cautiously")]
    pub remove: Option<bool>,
    /// compression
    #[arg(
        short = 'c',
        long = "compression",
        required = false,
        help = "value between 1-9"
    )]
    pub compression: Option<i64>,
    /// output filepath
    #[arg(
        short = 'o',
        long = "output",
        required = false,
        help = "output filepath for list of files"
    )]
    pub output: Option<PathBuf>,

    /// exclude expr
    #[arg(
        short = 'e',
        long = "exclude",
        required = false,
        help = "exclude by pattern"
    )]
    pub exclude: Option<String>,

    /// levels
    #[arg(
        short = 'l',
        long = "levels",
        required = false,
        help = "traversal depth"
    )]
    pub levels: Option<usize>,
    
    /// size cap in megabytes
    #[arg(
        short = 's',
        long = "size",
        required = false,
        help = "size cap in MB (ie for 1MB use 1 instead 1024 ) of a file for search"
    )]
    pub size: Option<u64>,
    
    /// inclusion
    #[arg(
        short = 'i',
        long = "include",
        required = false,
        help = "only include files with this pattern"
    )]
    pub include: Option<String>,
    
}

impl Context {
    pub fn new() -> Self {
        Self::parse()
    }
}
#[cfg(test)]
mod test_args {
    use super::*;

    #[test]
    fn test_default_values() {
        assert_eq!(Context::new().remove, None);
    }
}
