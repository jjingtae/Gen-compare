//! Strip transparent padding → square-pad → emit Tauri icon set.
//!
//! Usage: icon-gen <source.png> <out_dir>
//!
//! Writes: icon.ico, icon.png, 32x32.png, 128x128.png, 128x128@2x.png,
//!         Square30x30Logo.png, Square44x44Logo.png, Square71x71Logo.png,
//!         Square89x89Logo.png, Square107x107Logo.png, Square142x142Logo.png,
//!         Square150x150Logo.png, Square284x284Logo.png, Square310x310Logo.png,
//!         StoreLogo.png

use anyhow::Context;
use image::{imageops::FilterType, ImageBuffer, Rgba, RgbaImage};
use std::path::PathBuf;

fn main() -> anyhow::Result<()> {
    let src = std::env::args().nth(1).context("usage: icon-gen <src.png> <out_dir>")?;
    let out_dir = PathBuf::from(
        std::env::args().nth(2).context("usage: icon-gen <src.png> <out_dir>")?,
    );
    std::fs::create_dir_all(&out_dir)?;

    let dyn_img = image::open(&src).context("open source PNG")?;
    println!("source color type: {:?}", dyn_img.color());
    let img = dyn_img.to_rgba8();
    let (w, h) = img.dimensions();

    // Diagnostic: check whether the PNG has meaningful alpha. We sample the
    // four corner pixels — if the source really is transparent around a
    // centered icon (as Windows Photos viewer shows via its checkerboard),
    // those corners will be alpha=0. Otherwise the transparency was baked
    // into grey-checker pixels during a screenshot-save roundtrip.
    let corner_alphas = [
        img.get_pixel(0, 0)[3],
        img.get_pixel(w - 1, 0)[3],
        img.get_pixel(0, h - 1)[3],
        img.get_pixel(w - 1, h - 1)[3],
    ];
    println!("corner alphas: {:?}", corner_alphas);
    let has_real_alpha = corner_alphas.iter().any(|&a| a < 32);

    // 1) Find bounding box. Prefer alpha where present; otherwise fall back to
    //    detecting "not-background" pixels by comparing each pixel against the
    //    four corner samples — this is how we handle a flattened screenshot
    //    whose grey checkerboard looks like transparency but is real pixels.
    let (min_x, min_y, max_x, max_y) = if has_real_alpha {
        println!("using alpha-channel cropping");
        alpha_bbox(&img)
    } else {
        println!("no alpha — using corner-color cropping (screenshot flattened)");
        color_bbox(&img)
    };

    if min_x >= max_x || min_y >= max_y {
        anyhow::bail!("could not detect non-background region");
    }
    let cw = max_x - min_x + 1;
    let ch = max_y - min_y + 1;
    let mut cropped = image::imageops::crop_imm(&img, min_x, min_y, cw, ch).to_image();

    // If the source had a baked-in checkerboard background (no real alpha),
    // knock it out with a single hard threshold. The icon has crisp,
    // saturated colors so a clean cut-off preserves edges better than a
    // soft gradient (which was making the rounded corners look hazy).
    if !has_real_alpha {
        let inset = 2u32.min(w / 20).max(2);
        let corners = [
            *img.get_pixel(inset, inset),
            *img.get_pixel(w - 1 - inset, inset),
            *img.get_pixel(inset, h - 1 - inset),
            *img.get_pixel(w - 1 - inset, h - 1 - inset),
        ];
        let thresh: i32 = 22;
        for p in cropped.pixels_mut() {
            let is_bg = corners.iter().any(|c| {
                let dr = (p[0] as i32 - c[0] as i32).abs();
                let dg = (p[1] as i32 - c[1] as i32).abs();
                let db = (p[2] as i32 - c[2] as i32).abs();
                dr <= thresh && dg <= thresh && db <= thresh
            });
            if is_bg { p[3] = 0; }
        }
    }

    // 2) Pad to square, centered on transparent background.
    let side = cw.max(ch);
    let mut square: RgbaImage = ImageBuffer::from_pixel(side, side, Rgba([0, 0, 0, 0]));
    let ox = ((side - cw) / 2) as i64;
    let oy = ((side - ch) / 2) as i64;
    image::imageops::overlay(&mut square, &cropped, ox, oy);

    println!(
        "source: {}x{} -> cropped: {}x{} -> square: {}x{}",
        w, h, cw, ch, side, side
    );

    // 3) Emit resized PNGs.
    let resize = |target: u32| -> RgbaImage {
        image::imageops::resize(&square, target, target, FilterType::Lanczos3)
    };

    // Standard Tauri names.
    resize(32).save(out_dir.join("32x32.png"))?;
    resize(128).save(out_dir.join("128x128.png"))?;
    resize(256).save(out_dir.join("128x128@2x.png"))?;
    resize(512).save(out_dir.join("icon.png"))?;

    // Windows Store tile names (required when bundle targets MSI / Microsoft Store).
    for (name, s) in [
        ("Square30x30Logo.png", 30),
        ("Square44x44Logo.png", 44),
        ("Square71x71Logo.png", 71),
        ("Square89x89Logo.png", 89),
        ("Square107x107Logo.png", 107),
        ("Square142x142Logo.png", 142),
        ("Square150x150Logo.png", 150),
        ("Square284x284Logo.png", 284),
        ("Square310x310Logo.png", 310),
        ("StoreLogo.png", 50),
    ] {
        resize(s).save(out_dir.join(name))?;
    }

    // 4) Multi-resolution ICO.
    let mut ico_dir = ico::IconDir::new(ico::ResourceType::Icon);
    for s in [16u32, 24, 32, 48, 64, 128, 256] {
        let r = resize(s);
        let (rw, rh) = r.dimensions();
        let icon_img = ico::IconImage::from_rgba_data(rw, rh, r.into_raw());
        ico_dir.add_entry(ico::IconDirEntry::encode(&icon_img)?);
    }
    let mut f = std::fs::File::create(out_dir.join("icon.ico"))?;
    ico_dir.write(&mut f)?;

    println!("wrote icon set to {}", out_dir.display());
    Ok(())
}

/// Bounding box of pixels whose alpha is meaningfully opaque.
fn alpha_bbox(img: &RgbaImage) -> (u32, u32, u32, u32) {
    let (w, h) = img.dimensions();
    let (mut minx, mut miny, mut maxx, mut maxy) = (w, h, 0u32, 0u32);
    for y in 0..h {
        for x in 0..w {
            if img.get_pixel(x, y)[3] > 8 {
                if x < minx { minx = x; }
                if y < miny { miny = y; }
                if x > maxx { maxx = x; }
                if y > maxy { maxy = y; }
            }
        }
    }
    (minx, miny, maxx, maxy)
}

/// Bounding box for a flattened screenshot. Samples the four corners as
/// "background" colors — any pixel whose channel distance exceeds a threshold
/// against every corner sample is treated as foreground content. Designed for
/// a transparent-viewer screenshot where the checkerboard grid is baked in.
fn color_bbox(img: &RgbaImage) -> (u32, u32, u32, u32) {
    let (w, h) = img.dimensions();
    // Corner samples — taken a few pixels in so we avoid anti-aliased edges.
    let inset = 2u32.min(w / 20).max(2);
    let corners = [
        *img.get_pixel(inset, inset),
        *img.get_pixel(w - 1 - inset, inset),
        *img.get_pixel(inset, h - 1 - inset),
        *img.get_pixel(w - 1 - inset, h - 1 - inset),
    ];
    // A pixel is "background" if its RGB distance to any corner sample is
    // within `thresh`. Tuned for the ~20-value difference between the two
    // checker squares.
    let thresh: i32 = 24;
    let is_bg = |p: &Rgba<u8>| -> bool {
        corners.iter().any(|c| {
            let dr = p[0] as i32 - c[0] as i32;
            let dg = p[1] as i32 - c[1] as i32;
            let db = p[2] as i32 - c[2] as i32;
            dr.abs() <= thresh && dg.abs() <= thresh && db.abs() <= thresh
        })
    };
    let (mut minx, mut miny, mut maxx, mut maxy) = (w, h, 0u32, 0u32);
    for y in 0..h {
        for x in 0..w {
            if !is_bg(img.get_pixel(x, y)) {
                if x < minx { minx = x; }
                if y < miny { miny = y; }
                if x > maxx { maxx = x; }
                if y > maxy { maxy = y; }
            }
        }
    }
    (minx, miny, maxx, maxy)
}
