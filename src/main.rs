mod ctx;
mod search;
mod zip_dir;

use crate::search::determine_files_with_links;
use crate::zip_dir::{extract_dir, get_meta, set_meta, zip_dir, MetaApply};
use ctx::Context;
use std::io::{stdin, BufWriter, Write};
use std::os::windows::fs::MetadataExt;
use std::path::{Path, PathBuf};
use indicatif::{ProgressBar, ProgressStyle};
use walkdir::{DirEntry, WalkDir};

#[cfg(target_os = "windows")]
#[no_mangle]
pub fn handle(files: Vec<PathBuf>, compression_level: i64) {

    let pb = ProgressBar::new(files.len() as u64);
    pb.set_style(ProgressStyle::default_bar()
        .template("{msg}\n{bar:40} {pos}/{len}")
        .unwrap()
        .progress_chars("##-"));
    
    for file in files {

        let meta_apply: MetaApply = get_meta(file.clone());

        let temporary_directory_name = String::from("tmp_output_") + chrono::Local::now().timestamp().to_string().as_ref();
        let target = Path::join(
            file.parent().unwrap().to_str().unwrap().as_ref(),
            temporary_directory_name,
        );

        let target = target.to_str().expect("failed joining paths to target");

        match extract_dir(file.as_path(), target) {
            Ok(_) => match zip_dir(target, file.as_path(), compression_level) {
                Ok(_) => {
                    set_meta(file.as_path(), meta_apply);
                    pb.inc(1);
                }
                Err(e) => panic!("{e}"),
            },
            Err(e) => panic!("{e}"),
        }
    }
}

#[no_mangle]
pub fn output_list(list: &[PathBuf], f: &mut std::fs::File) {
    let mut buffered_writer = BufWriter::new(f);
    buffered_writer
        .write_all(b"path\n")
        .expect("Failed to Write to Designated Output File");
    for path in list.iter() {
        buffered_writer
            .write_all(path.as_os_str().as_encoded_bytes())
            .expect("Failed to Write to Designated Output File");
        buffered_writer
            .write_all(b"\n")
            .expect("Failed to Write to Designated Output File");
    }
}

#[no_mangle]
fn get_searchable_files(c: &Context) -> Vec<DirEntry> {
    if let Some(path) = c.dir.as_ref() {
        
        // even if no param is passed we still want to cap at a 1gb
        let max_file_size: u64 = c.size.unwrap_or(1073741824 / 1024) * 1024;
        
        let wd = {
            if let Some(depth) = c.levels {
                WalkDir::new(path).max_depth(depth)
            } else {
                WalkDir::new(path)
            }
        };
        
        wd.same_file_system(true).into_iter().filter_map(Result::ok).filter_map(|file| {
            // remove records with too large a file size
            if file.metadata().as_ref().unwrap().file_size() >= max_file_size {
                None
            } else if let Some(exclude_pattern) = c.exclude.as_ref() {
                // if the file DOES NOT (false) contain the pattern to be EXCLUDED include it
                if file.path().to_str().unwrap().to_lowercase().contains(exclude_pattern.to_lowercase().as_str()).eq(&false) {
                    if let Some(include) = c.include.as_ref() {
                        // if the file DOES (true) contain the pattern to be INCLUDED include it
                        if file.path().to_str().unwrap().to_lowercase().contains(include.to_lowercase().as_str()).eq(&true) {
                            Some(file) 
                        // the file DOES NOT contain what should be INCLUDED exclude it
                        } else {
                            None
                        } 
                    // no included clause provided so the result of the exclusion holds
                    } else {
                        Some(file)
                    }
                // failed exclusion clause no need to check inclusion
                } else {
                    None
                }
            // no Exclusion clause provided, still need to check for Inclusion clause
            } else if let Some(include) = c.include.as_ref() {
                // if the file DOES (true) contain the pattern to be INCLUDED include it
                if file.path().to_str().unwrap().to_lowercase().contains(include.to_lowercase().as_str()).eq(&true) {
                    Some(file)
                    // the file DOES NOT contain what should be INCLUDED exclude it
                } else {
                    None
                }
                // no included or excluded clause provided so the result of the exclusion holds
            } else {
                Some(file)
            }
        }).collect()
    } else {
        panic!("failed to get count from directory")
    }
}

/// delete definedNames from workbook.xml only those with a [1]
/// delete externalLinks from workbook.xml
/// delete externalLinks folder entirely
fn main() {

    let ctx = Context::new();
    let files = get_searchable_files(&ctx);
    let mut pb = ProgressBar::new(files.len() as u64);

    pb.set_style(ProgressStyle::default_bar()
        .template("{msg}\n{bar:40} {pos}/{len}")
        .unwrap()
        .progress_chars("##-"));
    
    // parse directories and get all necessary files
    let list = determine_files_with_links(files, &mut pb);
    
    // clear load
    pb.finish_and_clear();
    
    // check if user would like to remove links
    let remove = ctx.remove.unwrap_or(false);

    // generate output from list
    if let Some(write_to_path) = ctx.output {
        if let Ok(ref mut f) = std::fs::File::create(write_to_path) {
            output_list(list.as_ref(), f);
        } else {
            panic!("failed to create file at path");
        }
    } else {
        let default_file_name = String::from("workbook_links_found - ")
            + chrono::Local::now().timestamp().to_string().as_ref()
            + ".txt";
        if let Ok(ref mut f) = std::fs::File::create(default_file_name) {
            output_list(list.as_ref(), f);
        } else {
            panic!("failed to create file at path");
        }
    }

    // remove links if user has asked and then double check
    if remove {
        println!("are you sure you want to remove links: Y/N?");
        let mut buffer = String::new();
        let _ = stdin().read_line(&mut buffer).expect("failed to get input");
        if buffer.trim() == "Y" {
            handle(list, ctx.compression.unwrap_or(3))
        } else {
            println!("not removing. closing.")
        }
    }
}
