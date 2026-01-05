#![cfg_attr(not(debug_assertions), windows_subsystem = "windows")]

// Debug builds will use the console subsystem (so println! works); release builds will be GUI on Windows.
use anyhow::Result;
use chrono::Local;
use clap::Parser;
use core::f64;
use csv::ReaderBuilder;
use rust_xlsxwriter::{
    chart::{ChartAxisCrossing, ChartAxisTickType, ChartFont, ChartFormat, ChartLine, ChartType},
    workbook::Workbook,
    Chart,
};
use std::{
    fs::File,
    io::{BufRead, BufReader},
    path::{Path, PathBuf},
};

const DPI: u32 = 96;
const HEIGHT: u32 = 6 * DPI;
const WIDTH: u32 = 8 * DPI;

// Choose a 'nice' step size (0.1,0.2,0.5,1,2,5,10,20 × 10^n). Supports decimal steps
// and considers smaller magnitudes for values < 1.
fn nice_step(raw: f64) -> f64 {
    if raw <= 0.0 {
        return 1.0;
    }
    // Work in integer exponents but consider a wider range so we can return
    // reasonably small steps for tiny raw values.
    let exp = raw.abs().log10().floor() as i32;
    let multipliers = [0.01, 0.1, 0.2, 0.5, 1.0, 2.0, 5.0, 10.0, 20.0];
    for e in (exp - 2)..=(exp + 1) {
        let base = 10f64.powi(e);
        for &m in &multipliers {
            let s = m * base;
            if s >= raw {
                return s;
            }
        }
    }
    // Fallback
    20.0 * 10f64.powi(exp)
}

#[derive(Debug)]
struct BasicPlotData {
    x: Vec<f64>,
    y: Vec<f64>,
}

#[derive(Debug)]
struct DifferentialPlotData {
    header: Vec<String>,
    x_g: Vec<f64>,
    intensity: Vec<f64>,
    frequency: f64,
}

trait Plottable {
    fn from_file(path: &Path) -> Result<Self>
    where
        Self: Sized;
    fn to_excel(self, out_filepath: &Path, base_name: Option<&str>) -> Result<()>;
}

impl BasicPlotData {
    fn new() -> Self {
        return BasicPlotData {
            x: Vec::new(),
            y: Vec::new(),
        };
    }
}

impl Plottable for BasicPlotData {
    fn from_file(path: &Path) -> Result<Self>
    where
        Self: Sized,
    {
        let file = File::open(path)?;
        let mut res = BasicPlotData::new();

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

    fn to_excel(self, out_filepath: &Path, base_name: Option<&str>) -> Result<()> {
        let end = self.x.len() + 1;
        let mut workbook = Workbook::new();
        let worksheet = workbook.add_worksheet();

        worksheet.write(0, 0, "x")?;
        worksheet.write(0, 1, "y")?;
        worksheet.write_column(1, 0, self.x.clone())?;
        worksheet.write_column(1, 1, self.y.clone())?;

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
        chart.set_height(HEIGHT);
        chart.set_width(WIDTH);

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

        chart.legend().set_hidden();

        // Insert chart at D2
        worksheet.insert_chart(1, 3, &chart)?;

        workbook.save(out_filepath)?;

        Ok(())
    }
}

impl Plottable for DifferentialPlotData {
    fn from_file(path: &Path) -> Result<Self>
    where
        Self: Sized,
    {
        let file = File::open(path)?;
        let buf = BufReader::new(file);
        let mut lines = buf.lines();

        // Collect header lines until we find the data header starting with "X [G]"
        let mut header: Vec<String> = Vec::new();
        loop {
            match lines.next() {
                Some(line_res) => {
                    let line = line_res?;
                    if line.trim_start().starts_with("X [G]") {
                        break;
                    } else {
                        header.push(line);
                    }
                }
                None => {
                    return Err(anyhow::anyhow!(
                        "Reached EOF before finding data header 'X [G]'."
                    ));
                }
            }
        }

        // If the user cancels, user_value remains None. If they enter invalid text, show an
        // error and re-prompt until they cancel or provide a valid float.
        let frequency: f64 = loop {
            match tinyfiledialogs::input_box(
                "Enter Frequency",
                "Enter a floating-point value (or click Cancel):",
                "1.0",
            ) {
                Some(s) => match s.trim().parse::<f64>() {
                    Ok(v) => break v,
                    Err(_) => {
                        tinyfiledialogs::message_box_ok(
                            "Invalid input",
                            "Please enter a valid floating point number.",
                            tinyfiledialogs::MessageBoxIcon::Error,
                        );
                        continue;
                    }
                },
                None => {
                    panic!("User cancelled input dialog.");
                }
            }
        };

        // Parse remaining lines as data pairs: x_g and intensity
        let mut x_g: Vec<f64> = Vec::new();
        let mut intensity: Vec<f64> = Vec::new();

        for line_res in lines {
            let line = line_res?;
            let s = line.trim();
            if s.is_empty() {
                continue;
            }

            let mut parts = s.split_whitespace();
            let a = parts.next().ok_or(anyhow::anyhow!("Missing X [G] value"))?;
            let b = parts
                .next()
                .ok_or(anyhow::anyhow!("Missing intensity value"))?;
            x_g.push(a.parse()?);
            intensity.push(b.parse()?);
        }

        if x_g.len() != intensity.len() {
            return Err(anyhow::anyhow!(
                "Malformed input, 'X [G]' and 'Intensity' must be equal length"
            ));
        }

        // let mut paired: Vec<_> = x_g.into_iter().zip(intensity.into_iter()).collect();
        // paired.sort_by(|(a1, _), (a2, _)| a1.partial_cmp(a2).unwrap());

        // let (x_g_sorted, intensity_sorted): (Vec<f64>, Vec<f64>) = paired.into_iter().unzip();

        Ok(DifferentialPlotData {
            header,
            // x_g: x_g_sorted,
            // intensity: intensity_sorted,
            x_g,
            intensity,
            frequency,
        })
    }

    fn to_excel(self, out_filepath: &Path, base_name: Option<&str>) -> Result<()> {
        let mut workbook = Workbook::new();
        let worksheet = workbook.add_worksheet();
        for (i, line) in self.header.iter().enumerate() {
            worksheet.write(i as u32, 0, line)?;
        }

        // Write column headers after the file header
        let header_len = self.header.len() as u32;
        worksheet.write(header_len, 0, "X [G]")?;
        worksheet.write(header_len, 1, "Intensity")?;
        worksheet.write(header_len, 2, "g-factor")?;
        worksheet.write(header_len, 3, "Intensity")?;

        // Data start (zero-based row index)
        let data_start = header_len + 1;
        let data_end = data_start + (self.x_g.len() as u32) - 1;

        // Excel rows are 1-based in formulas
        let excel_data_start = data_start + 1;
        let excel_data_end = data_end + 1;

        worksheet.write_column(data_start, 0, self.x_g.clone())?;
        worksheet.write_column(data_start, 1, self.intensity.clone())?;

        // Write g-factor formula in column C for each row: (714.8 * frequency) / A<row>
        for (i, _) in self.x_g.iter().enumerate() {
            let row_idx = data_start + i as u32; // 0-based for write methods
            let excel_row_num = excel_data_start + i as u32; // 1-based for formula reference
            let formula = format!("=(714.8*{})/A{}", self.frequency, excel_row_num);
            worksheet.write_formula(row_idx, 2, &*formula)?;
        }

        // Write intensity formula in column D for each row: B<row> / 1000
        for (i, _) in self.x_g.iter().enumerate() {
            let row_idx = data_start + i as u32; // 0-based for write methods
            let excel_row_num = excel_data_start + i as u32; // 1-based for formula reference
            let formula = format!("=B{}/1000", excel_row_num);
            worksheet.write_formula(row_idx, 3, &*formula)?;
        }

        /* ------------------ First chart -------------------------- */

        let mut chart_a = Chart::new(ChartType::ScatterSmooth);
        let chart_title = match base_name {
            None => Local::now().format("%Y-%m-%d_%H-%M-%S").to_string(),
            Some(x) => x.to_string(),
        };

        let cat_range = format!("=Sheet1!$A${}:$A${}", excel_data_start, excel_data_end);
        let val_range = format!("=Sheet1!$B${}:$B${}", excel_data_start, excel_data_end);
        chart_a
            .add_series()
            .set_categories(&cat_range)
            .set_values(&val_range)
            .set_format(
                ChartFormat::new().set_line(ChartLine::new().set_color("#00008B").set_width(0.5)),
            );

        chart_a
            .title()
            .set_name(&chart_title)
            .set_font(ChartFont::new().set_size(14));
        chart_a.set_height(HEIGHT);
        chart_a.set_width(WIDTH);

        let mut axis_font = ChartFont::new();
        axis_font.set_size(9);

        let target_ticks = 8.0;

        // X axis: snap min/max outward to multiples of a nice step
        let x_g_min_raw = self.x_g.iter().cloned().reduce(f64::min).unwrap_or(0.0);
        let x_g_max_raw = self.x_g.iter().cloned().reduce(f64::max).unwrap_or(0.0);

        // Simpler policy: force X major ticks to 20 with 10 minor ticks between.
        // Snap bounds to multiples of 20 so gridlines align cleanly. Use the next
        // 20-step as the start (ceil) — it's acceptable to cut one left step for
        // a tidier axis.
        let x_step = 20.0_f64;
        let mut x_g_min = (x_g_min_raw / x_step).ceil() * x_step; // start at next multiple (may cut leftmost step)
        let mut x_g_max = (x_g_max_raw / x_step).ceil() * x_step;
        if (x_g_max - x_g_min).abs() < std::f64::EPSILON {
            x_g_max = x_g_min + x_step;
            x_g_min = x_g_max - x_step; // ensure at least one step range
        }
        // Ensure data are included on the right; if snapping somehow excluded the data max, expand minimally
        if x_g_max < x_g_max_raw {
            x_g_max = x_g_min + x_step * (((x_g_max_raw - x_g_min) / x_step).ceil().max(1.0));
        }

        // Compute symmetric Y axis around zero so 0 is centered, while ensuring the
        // full intensity range is visible. Add a small padding (5%). If the data
        // are all zero, fall back to a default range of [-1, 1]. Then snap the
        // Y limits to a 'nice' step so major gridlines divide cleanly.
        let intensity_min_raw = self
            .intensity
            .iter()
            .cloned()
            .reduce(f64::min)
            .unwrap_or(0.0);
        let intensity_max_raw = self
            .intensity
            .iter()
            .cloned()
            .reduce(f64::max)
            .unwrap_or(0.0);
        let max_abs = intensity_min_raw.abs().max(intensity_max_raw.abs());
        let y_limit = if max_abs == 0.0 {
            1.0
        } else {
            // 5% padding
            max_abs * 1.05
        };

        let raw_y_step = (2.0 * y_limit) / target_ticks;
        let y_step = nice_step(raw_y_step);

        // Snap y limits to integer multiples of step and keep symmetrical about 0
        let steps_needed = (y_limit / y_step).ceil();
        let intensity_max = steps_needed * y_step;
        let intensity_min = -intensity_max;

        let x_crossing = ChartAxisCrossing::Min;
        chart_a
            .x_axis()
            .set_name("Field (G)")
            .set_name_font(&axis_font)
            .set_format(ChartFormat::new().set_line(ChartLine::new().set_color("#000000")))
            .set_min(x_g_min)
            .set_max(x_g_max)
            .set_major_unit(x_step)
            .set_minor_unit(x_step / 5.0)
            .set_major_tick_type(ChartAxisTickType::Inside)
            .set_minor_tick_type(ChartAxisTickType::Inside)
            .set_major_gridlines(false)
            .set_minor_gridlines(false)
            .set_minor_gridlines(false)
            .set_crossing(x_crossing);

        let y_crossing = ChartAxisCrossing::Min;
        chart_a
            .y_axis()
            .set_name("Intensity (a. u.)")
            .set_name_font(&axis_font)
            .set_format(ChartFormat::new().set_line(ChartLine::new().set_color("#000000")))
            .set_min(intensity_min)
            .set_max(intensity_max)
            .set_major_unit(y_step)
            .set_major_tick_type(ChartAxisTickType::Inside)
            .set_minor_tick_type(ChartAxisTickType::Inside)
            .set_major_gridlines(false)
            .set_minor_gridlines(false)
            .set_crossing(y_crossing);

        chart_a.legend().set_hidden();

        /* ------------------ Second chart -------------------------- */

        // Compute the columns for g-factor and normalized intensity
        // We don't write these values (excel calculates on the fly with the formulas)
        // But we still need them for the ranges (axis) in the chart.

        let g_factor = self
            .x_g
            .iter()
            .map(|&x| (714.8 * self.frequency) / x)
            .collect::<Vec<f64>>();
        let normalized_intensity = self
            .intensity
            .iter()
            .map(|&x| x / 1000.0)
            .collect::<Vec<f64>>();

        let mut chart_b = Chart::new(ChartType::ScatterSmooth);
        let chart_title = match base_name {
            None => Local::now().format("%Y-%m-%d_%H-%M-%S").to_string(),
            Some(x) => x.to_string(),
        };

        let cat_range = format!("=Sheet1!$C${}:$C${}", excel_data_start, excel_data_end);
        let val_range = format!("=Sheet1!$D${}:$D${}", excel_data_start, excel_data_end);
        chart_b
            .add_series()
            .set_categories(&cat_range)
            .set_values(&val_range)
            .set_format(
                ChartFormat::new().set_line(ChartLine::new().set_color("#00008B").set_width(0.5)),
            );

        chart_b
            .title()
            .set_name(&chart_title)
            .set_font(ChartFont::new().set_size(14));
        chart_b.set_height(HEIGHT);
        chart_b.set_width(WIDTH);

        let mut axis_font = ChartFont::new();
        axis_font.set_size(9);

        let target_ticks = 8.0;

        // X axis: snap min/max outward to multiples of a nice step
        let g_factor_min_raw = g_factor.iter().cloned().reduce(f64::min).unwrap_or(0.0);
        let g_factor_max_raw = g_factor.iter().cloned().reduce(f64::max).unwrap_or(0.0);

        let x_step = 0.01_f64;

        let g_factor_min = (g_factor_min_raw / x_step).floor() * x_step;
        let mut g_factor_max = (g_factor_max_raw / x_step).floor() * x_step;

        if (g_factor_max - g_factor_min).abs() < std::f64::EPSILON {
            g_factor_max = g_factor_min + x_step;
        }
        if g_factor_max <= g_factor_min {
            g_factor_max = g_factor_min + x_step;
        }

        // Compute symmetric Y axis around zero so 0 is centered, while ensuring the
        // full normalized_intensity range is visible. Add a small padding (5%). If the data
        // are all zero, fall back to a default range of [-1, 1]. Then snap the
        // Y limits to a 'nice' step so major gridlines divide cleanly.
        let normalized_intensity_min_raw = normalized_intensity
            .iter()
            .cloned()
            .reduce(f64::min)
            .unwrap_or(0.0);
        let normalized_intensity_max_raw = normalized_intensity
            .iter()
            .cloned()
            .reduce(f64::max)
            .unwrap_or(0.0);
        let max_abs = normalized_intensity_min_raw
            .abs()
            .max(normalized_intensity_max_raw.abs());
        let y_limit = if max_abs == 0.0 {
            1.0
        } else {
            // 5% padding
            max_abs * 1.05
        };

        let raw_y_step = (2.0 * y_limit) / target_ticks;
        let y_step = nice_step(raw_y_step);

        // Snap y limits to integer multiples of step and keep symmetrical about 0
        let steps_needed = (y_limit / y_step).ceil();
        let normalized_intensity_max = steps_needed * y_step;
        let normalized_intensity_min = -normalized_intensity_max;

        let x_crossing = ChartAxisCrossing::Max;
        chart_b
            .x_axis()
            .set_name("g value")
            .set_name_font(&axis_font)
            .set_format(ChartFormat::new().set_line(ChartLine::new().set_color("#000000")))
            .set_min(g_factor_min)
            .set_max(g_factor_max)
            .set_major_unit(x_step)
            .set_minor_unit(x_step / 5.0)
            .set_major_tick_type(ChartAxisTickType::Inside)
            .set_minor_tick_type(ChartAxisTickType::Inside)
            .set_major_gridlines(false)
            .set_minor_gridlines(false)
            .set_minor_gridlines(false)
            .set_crossing(x_crossing)
            .set_reverse();

        let y_crossing = ChartAxisCrossing::Min;
        chart_b
            .y_axis()
            .set_name("Intensity (a. u.)")
            .set_name_font(&axis_font)
            .set_format(ChartFormat::new().set_line(ChartLine::new().set_color("#000000")))
            .set_min(normalized_intensity_min)
            .set_max(normalized_intensity_max)
            .set_major_unit(y_step)
            .set_major_tick_type(ChartAxisTickType::Inside)
            .set_minor_tick_type(ChartAxisTickType::Inside)
            .set_major_gridlines(false)
            .set_minor_gridlines(false)
            .set_crossing(y_crossing);

        chart_b.legend().set_hidden();

        worksheet.insert_chart(3, 5, &chart_a)?;
        worksheet.insert_chart(3, 18, &chart_b)?;

        workbook.save(out_filepath)?;

        Ok(())
    }
}

#[derive(Debug)]
enum Plot {
    Basic(BasicPlotData),
    Differential(DifferentialPlotData),
}

fn read_input(path: &Path) -> Result<Plot> {
    let extension = path
        .extension()
        .and_then(|s| s.to_str())
        .unwrap_or(".xy")
        .to_lowercase();

    return match extension.as_str() {
        "asc" => Ok(Plot::Differential(DifferentialPlotData::from_file(path)?)),
        _ => Ok(Plot::Basic(BasicPlotData::from_file(path)?)),
    };
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

    if args.output.is_none() {};

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

    let data = read_input(&args.infile)?;
    match data {
        Plot::Basic(d) => d.to_excel(&unique_out_filepath, title.as_deref())?,
        Plot::Differential(d) => d.to_excel(&unique_out_filepath, title.as_deref())?,
    }

    opener::open(unique_out_filepath)?;

    Ok(())
}
