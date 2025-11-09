#![windows_subsystem = "windows"]

use anyhow::Result;
use chrono::Local;
use core::f64;
use csv::ReaderBuilder;
use rust_xlsxwriter::{
    chart::{ChartAxisTickType, ChartFont, ChartFormat, ChartLine, ChartType},
    workbook::Workbook,
    Chart,
};
use std::{fs::File, io::BufReader, path::{Path, PathBuf}};
use clap::Parser;

#[derive(Debug)]
struct PlotData {
    x: Vec<f64>,
    y: Vec<f64>,
}

impl PlotData {
    fn new() -> Self {
        return PlotData {
            x: Vec::new(),
            y: Vec::new(),
        };
    }

    fn to_excel(self, out_filepath: &Path, base_name: Option<&str>) -> Result<()> {
        let end = self.x.len() + 1;
        let mut workbook = Workbook::new();
        let worksheet = workbook.add_worksheet();

        worksheet.write(0, 0, "x")?;
        worksheet.write(0, 1, "y")?;
        worksheet.write_column(1, 0, self.x.clone())?;
        worksheet.write_column(1, 1, self.y.clone())?;

        let dpi = 96;
        let height = 6 * dpi;
        let width = 8 * dpi;

        let mut chart = Chart::new(ChartType::ScatterSmooth);
        let chart_title = match base_name {
            None => Local::now().format("%Y-%m-%d_%H-%M-%S").to_string(),
            Some(x) => x.to_string(),
        };

        let cat_range = format!("=Sheet1!$A$2:$A${}", end);
        let val_range = format!("=Sheet1!$B$2:$B${}", end);
        chart
            .add_series()
            .set_categories(&cat_range)
            .set_values(&val_range)
            .set_format(
                ChartFormat::new().set_line(ChartLine::new().set_color("#FF0000").set_width(0.5)),
            );
        chart
            .title()
            .set_name(&chart_title)
            .set_font(ChartFont::new().set_size(14));
        chart.set_height(height);
        chart.set_width(width);

        let mut axis_font = ChartFont::new();
        axis_font.set_size(9).unset_bold();

        let x_min = self
            .x
            .iter()
            .cloned()
            .reduce(f64::min)
            .unwrap_or(0.0)
            .floor();
        let x_max = self
            .x
            .iter()
            .cloned()
            .reduce(f64::max)
            .unwrap_or(0.0)
            .ceil();
        let y_min = self
            .y
            .iter()
            .cloned()
            .reduce(f64::min)
            .unwrap_or(0.0)
            .floor();
        let y_max = self
            .y
            .iter()
            .cloned()
            .reduce(f64::max)
            .unwrap_or(0.0)
            .ceil();

        chart
            .x_axis()
            .set_name("2Theta")
            .set_name_font(&axis_font)
            .set_format(ChartFormat::new().set_line(ChartLine::new().set_color("#000000")))
            .set_min(x_min)
            .set_max(x_max)
            .set_major_tick_type(ChartAxisTickType::Inside)
            .set_minor_tick_type(ChartAxisTickType::Inside)
            .set_major_gridlines(false)
            .set_minor_gridlines(false);

        chart
            .y_axis()
            .set_name("Counts")
            .set_name_font(&axis_font)
            .set_format(ChartFormat::new().set_line(ChartLine::new().set_color("#000000")))
            .set_min(y_min)
            .set_max(y_max)
            .set_major_tick_type(ChartAxisTickType::Inside)
            .set_minor_tick_type(ChartAxisTickType::Inside)
            .set_major_gridlines(false)
            .set_minor_gridlines(false);

        // Insert chart at D2
        worksheet.insert_chart(1, 3, &chart)?;

        workbook.save(out_filepath)?;

        Ok(())
    }
}

fn csv_read(path: &Path) -> Result<PlotData> {
    let file = File::open(path)?;
    let mut res = PlotData::new();
    let mut reader = ReaderBuilder::new()
        .delimiter(b' ')
        .from_reader(BufReader::new(file));

    for result in reader.records() {
        let record = result?;
        let (x, y) = (
            record.get(0).ok_or(anyhow::anyhow!("Missing x value"))?,
            record.get(1).ok_or(anyhow::anyhow!("Missing y value"))?,
        );
        res.x.push(x.parse()?);
        res.y.push(y.parse()?);
    }

    if res.x.len() != res.y.len() {
        return Err(anyhow::anyhow!(
            "Malformed input, X and Y must be equal length"
        ));
    }

    Ok(res)
}

fn unique_path(path: PathBuf) -> PathBuf {
    if !path.exists() {
        return path;
    }
    let mut count = 1;

    let parent = path.parent().unwrap_or_else(|| Path::new(""));
    let ext = path.extension().and_then(|s| s.to_str()).unwrap_or("xlsx");
    let stem = path
        .file_stem()
        .and_then(|s| s.to_str())
        .unwrap_or("_default_output")
        .to_string();

    let mut unique_path = path.clone();
    while unique_path.exists() {
        let new_filename = if ext.is_empty() {
            format!("{}_{}", stem, count)
        } else {
            format!("{}_{}.{}", stem, count, ext)
        };

        unique_path = parent.join(new_filename);
        count += 1;
    }

    unique_path
}

#[derive(Parser, Debug)]
#[command(author, version, about)]
struct CliArgs {
    /// Input file
    infile: PathBuf,

    /// Optional chart title. Default: Infer based on input filename.
    #[arg(short = 't', long = "title")]
    title: Option<String>,

    /// Output filename. Default: Infer based on input filename.
    #[arg(short = 'o', long = "output")]
    output: Option<PathBuf>,
}

fn main() -> Result<()> {
    let args = CliArgs::parse();

    let in_file = &args.infile;
    let base_name = in_file
        .file_stem()
        .and_then(|s| s.to_str())
        .unwrap_or(&in_file.to_string_lossy())
        .to_string();

    if args.output.is_none() {
    };

    let unique_out_filepath = match args.output {
        None => {
            let out_filename = format!("{}.xlsx", base_name);
            let out_filepath = match in_file.parent() {
                None => PathBuf::from(out_filename),
                Some(x) => x.join(out_filename),
            };
            unique_path(out_filepath)
        }
        Some(x) => x,
    };

    let title = match args.title {
        None => Some(base_name),
        Some(x) => Some(x),
    };

    let data = csv_read(&args.infile)?;
    data.to_excel(&unique_out_filepath, title.as_deref())?;

    Ok(())
}
