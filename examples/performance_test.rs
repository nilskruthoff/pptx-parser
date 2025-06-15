//! Performance benchmark example for the pptx-to-md crate
//!
//! This example measures and compares the performance of different parsing approaches:
//! - Single-threaded parsing
//! - Single-threading wih streaming
//! - Optimized multithreaded parsing with Rayon
//!
//! It also provides timing for individual operations to identify bottlenecks.
//!
//! Run with: cargo run --release --example performance_test <path/to/your/presentation.pptx> [iterations]

use pptx_to_md::{ParserConfig, PptxContainer, Result};
use rayon::prelude::*;
use std::env;
use std::path::Path;
use std::time::{Duration, Instant};

struct Benchmark {
    name: String,
    start_time: Instant,
    results: Vec<Duration>,
}

impl Benchmark {
    fn new(name: &str) -> Self {
        println!("Starting benchmark: {}", name);
        Benchmark {
            name: name.to_string(),
            start_time: Instant::now(),
            results: Vec::new(),
        }
    }

    fn measure<F, T>(&mut self, mut f: F) -> T
    where
        F: FnMut() -> T,
    {
        let start = Instant::now();
        let result = f();
        let duration = start.elapsed();
        self.results.push(duration);
        println!("  Operation took: {:?}", duration);
        result
    }

    fn report(&self) {
        if self.results.is_empty() {
            println!("No measurements for {}", self.name);
            return;
        }

        let total = self.start_time.elapsed();
        let count = self.results.len();
        let sum: Duration = self.results.iter().sum();
        let avg = sum / count as u32;
        let min = self.results.iter().min().unwrap();
        let max = self.results.iter().max().unwrap();

        println!("\nBenchmark Results for {}", self.name);
        println!("----------------------------");
        println!("Total time: {:?}", total);
        println!("Operations: {}", count);
        println!("Average time per operation: {:?}", avg);
        println!("Min time: {:?}", min);
        println!("Max time: {:?}", max);
        println!("----------------------------\n");
    }
}

fn main() -> Result<()> {
    // Get the PPTX file path and optional iteration count from command line arguments
    let args: Vec<String> = env::args().collect();
    let pptx_path = if args.len() > 1 {
        &args[1]
    } else {
        eprintln!("Usage: cargo run --example performance_test <path/to/presentation.pptx> [iterations]");
        return Ok(());
    };

    let iterations = if args.len() > 2 {
        args[2].parse().unwrap_or(5)
    } else {
        10 // Default to 10 iterations
    };

    println!("Performance testing with {} iterations on: {}", iterations, pptx_path);

    
    
    // =========== Single-threaded Approach ===========
    let mut single_thread_bench = Benchmark::new("Single-threaded parsing");

    let mut total_slides = 0;

    for i in 0..iterations {
        println!("\nIteration {} (Single-threaded)", i + 1);

        // Measure container creation
        let mut container = single_thread_bench.measure(|| {
            let config = ParserConfig::builder()
                .extract_images(true)
                .build();
            PptxContainer::open(Path::new(pptx_path), config).expect("Failed to open PPTX")
        });

        println!("  Found {} slides in the presentation", container.slide_count);

        // Measure parsing
        let slides = single_thread_bench.measure(|| {
            container.parse_all().expect("Failed to parse slides")
        });

        // Measure conversion
        let _md_content = single_thread_bench.measure(|| {
            slides.iter()
                .filter_map(|slide| slide.convert_to_md())
                .collect::<Vec<String>>()
        });

        total_slides += slides.len();
    }

    single_thread_bench.report();
    println!("Average slides per presentation: {}", total_slides / iterations);



    // =========== Single-threaded Streamed Approach ===========
    let mut single_thread_streamed_bench = Benchmark::new("Single-threaded streamed parsing");

    total_slides = 0;

    for i in 0..iterations {
        println!("\nIteration {} (Single-threaded streamed)", i + 1);

        // Measure container creation
        let mut container = single_thread_streamed_bench.measure(|| {
            let config = ParserConfig::builder()
                .extract_images(true)
                .build();
            PptxContainer::open(Path::new(pptx_path), config).expect("Failed to open PPTX")
        });

        println!("  Found {} slides in the presentation", container.slide_count);

        // Zähle die Slides im Voraus für die statistische Auswertung
        let expected_slides = container.slide_count;

        // Measure slide processing (including parsing and conversion)
        let slides_processed = single_thread_streamed_bench.measure(|| {
            let mut processed = 0;

            // Process slides one by one using the iterator
            for slide_result in container.iter_slides() {
                match slide_result {
                    Ok(slide) => {
                        // Konvertiere den Slide zu Markdown
                        let _md_content = slide.convert_to_md();
                        processed += 1;
                    },
                    Err(e) => {
                        eprintln!("Error processing slide: {:?}", e);
                    }
                }
            }

            processed
        });

        println!("  Processed {} out of {} slides", slides_processed, expected_slides);
        total_slides += slides_processed;
    }

    single_thread_streamed_bench.report();
    println!("Average slides per presentation: {}", total_slides / iterations);



    // =========== Optimized Multi-threaded Approach ===========
    let mut optimized_multi_thread_bench = Benchmark::new("Optimized Multi-threaded parsing");

    total_slides = 0;

    for i in 0..iterations {
        println!("\nIteration {} (Optimized Multi-threaded)", i + 1);

        // Container öffnen mit der gewünschten Konfiguration
        let mut container = optimized_multi_thread_bench.measure(|| {
            let config = ParserConfig::builder()
                .extract_images(true)
                .build();
            PptxContainer::open(Path::new(pptx_path), config).expect("Failed to open PPTX")
        });

        println!("  Found {} slides in the presentation", container.slide_count);

        // Verwende die neue optimierte Multi-Threading-Methode
        let slides = optimized_multi_thread_bench.measure(|| {
            container.parse_all_multi_threaded().expect("Failed to parse slides")
        });

        println!("  Successfully processed {} slides", slides.len());

        // Parallel zu Markdown konvertieren (bleibt unverändert)
        let _md_content = optimized_multi_thread_bench.measure(|| {
            slides.par_iter()
                .filter_map(|slide| slide.convert_to_md())
                .collect::<Vec<String>>()
        });

        total_slides += slides.len();
    }

    optimized_multi_thread_bench.report();
    println!("Average slides per presentation: {}", total_slides / iterations);

    // =========== Performance Comparison ===========
    if !single_thread_bench.results.is_empty() &&
        !single_thread_streamed_bench.results.is_empty() &&
        !optimized_multi_thread_bench.results.is_empty() {

        let single_avg: Duration = single_thread_bench.results.iter().sum::<Duration>() /
            single_thread_bench.results.len() as u32;
        let single_streamed_avg: Duration = single_thread_streamed_bench.results.iter().sum::<Duration>() /
            single_thread_streamed_bench.results.len() as u32;
        let optimized_multi_avg: Duration = optimized_multi_thread_bench.results.iter().sum::<Duration>() /
            optimized_multi_thread_bench.results.len() as u32;

        println!("\nPerformance Comparison");
        println!("=====================");
        println!("Single-threaded average: {:?}", single_avg);
        println!("Single-threaded streaming average: {:?}", single_streamed_avg);
        println!("Optimized multi-threaded average: {:?}", optimized_multi_avg);

        // Compare single-threaded vs single-threaded streaming
        if single_avg > single_streamed_avg {
            let speedup = single_avg.as_secs_f64() / single_streamed_avg.as_secs_f64();
            println!("Single-threaded streaming is {:.2}x faster than single-threaded", speedup);
        } else {
            let slowdown = single_streamed_avg.as_secs_f64() / single_avg.as_secs_f64();
            println!("Single-threaded streaming is {:.2}x slower than single-threaded", slowdown);
        }

        // Compare single-threaded vs optimized multithreaded
        if single_avg > optimized_multi_avg {
            let speedup = single_avg.as_secs_f64() / optimized_multi_avg.as_secs_f64();
            println!("Optimized multi-threaded is {:.2}x faster than single-threaded", speedup);
        } else {
            let slowdown = optimized_multi_avg.as_secs_f64() / single_avg.as_secs_f64();
            println!("Optimized multi-threaded is {:.2}x slower than single-threaded", slowdown);
        }

        // Compare single-threaded streaming vs optimized multithreaded
        if single_streamed_avg > optimized_multi_avg {
            let speedup = single_streamed_avg.as_secs_f64() / optimized_multi_avg.as_secs_f64();
            println!("Optimized multi-threaded is {:.2}x faster than single-threaded streaming", speedup);
        } else {
            let slowdown = optimized_multi_avg.as_secs_f64() / single_streamed_avg.as_secs_f64();
            println!("Optimized multi-threaded is {:.2}x slower than single-threaded streaming", slowdown);
        }

        // Determine the overall fastest approach
        let fastest_approach = if single_avg <= single_streamed_avg && single_avg <= optimized_multi_avg {
            "Single-threaded"
        } else if single_streamed_avg <= single_avg && single_streamed_avg <= optimized_multi_avg {
            "Single-threaded streaming"
        } else {
            "Optimized multi-threaded"
        };

        println!("\nOverall result: {} approach is the fastest for this workload.", fastest_approach);
    }

    Ok(())
}