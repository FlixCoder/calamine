#![feature(test)]

extern crate test;

use calamine::{open_workbook, open_workbook_auto, Ods, Reader, Xls, Xlsb, Xlsx};
use std::fs::File;
use std::io::BufReader;
use test::Bencher;

fn read_big(path: &str) -> usize {
    let path = format!("{}/{}", env!("CARGO_MANIFEST_DIR"), path);
    let mut excel = open_workbook_auto(path).expect("cannot open excel file");

    let sheets = excel.sheet_names();
    let mut count = 0;
    for sheet in sheets {
        count += excel
            .worksheet_range(&sheet)
            .unwrap()
            .rows()
            .flat_map(|row| row.iter())
            .count();
    }
    count
}

#[bench]
fn bench_big_xls(b: &mut Bencher) {
    b.iter(|| read_big("tests/adhocallbabynames1996to2016.xls"));
}

fn count<R: Reader<BufReader<File>>>(path: &str) -> usize {
    let path = format!("{}/{}", env!("CARGO_MANIFEST_DIR"), path);
    let mut excel: R = open_workbook(&path).expect("cannot open excel file");

    let sheets = excel.sheet_names().to_owned();
    let mut count = 0;
    for s in sheets {
        count += excel
            .worksheet_range(&s)
            .unwrap()
            .rows()
            .flat_map(|r| r.iter())
            .count();
    }
    count
}

#[bench]
fn bench_xls(b: &mut Bencher) {
    b.iter(|| count::<Xls<_>>("tests/issues.xls"));
}

#[bench]
fn bench_xlsx(b: &mut Bencher) {
    b.iter(|| count::<Xlsx<_>>("tests/issues.xlsx"));
}

#[bench]
fn bench_xlsb(b: &mut Bencher) {
    b.iter(|| count::<Xlsb<_>>("tests/issues.xlsb"));
}

#[bench]
fn bench_ods(b: &mut Bencher) {
    b.iter(|| count::<Ods<_>>("tests/issues.ods"));
}

#[bench]
fn bench_xlsx_cells_reader(b: &mut Bencher) {
    fn count<R: Reader<BufReader<File>>>(path: &str) -> usize {
        let path = format!("{}/{}", env!("CARGO_MANIFEST_DIR"), path);
        let mut excel: Xlsx<_> = open_workbook(&path).expect("cannot open excel file");

        let sheets = excel.sheet_names().to_owned();
        let mut count = 0;
        for s in sheets {
            let mut cells_reader = excel.worksheet_cells_reader(&s).unwrap();
            while let Some(_) = cells_reader.next_cell().unwrap() {
                count += 1;
            }
        }
        count
    }
    b.iter(|| count::<Xlsx<_>>("tests/issues.xlsx"));
}

#[bench]
fn bench_xlsb_cells_reader(b: &mut Bencher) {
    fn count<R: Reader<BufReader<File>>>(path: &str) -> usize {
        let path = format!("{}/{}", env!("CARGO_MANIFEST_DIR"), path);
        let mut excel: Xlsb<_> = open_workbook(&path).expect("cannot open excel file");

        let sheets = excel.sheet_names().to_owned();
        let mut count = 0;
        for s in sheets {
            let mut cells_reader = excel.worksheet_cells_reader(&s).unwrap();
            while let Some(_) = cells_reader.next_cell().unwrap() {
                count += 1;
            }
        }
        count
    }
    b.iter(|| count::<Xlsx<_>>("tests/issues.xlsb"));
}
