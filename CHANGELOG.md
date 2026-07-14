# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [0.5.0] - 2026-07-14

### Added

- ODP support through `PresentationContainer`, including format detection and presentation metadata
- Hyperlinks, optional speaker notes and comments, and context-aware Markdown escaping

### Changed

- Placeholder positions now inherit from slide layouts and master slides for stable element ordering
- Test suite reorganized into unit and integration tests with dedicated fixtures

### Breaking

- Renamed `ParserConfig::include_slide_comment` to `include_slide_number_as_comment`; public struct literals must include the new `Run`, `Slide`, and `ParserConfig` fields
- `Slide::new` now accepts speaker notes and comments; `PptxContainer::get_full_image_path` was replaced by `resolve_target_path`

---

## [0.4.0] - 2025-07-01

### Added

- All slide elements now have a second `ElementPosition` parameter to save vertical and horizontal hierarchy
- `include_slide_comments` parameter to the `ParserConfig` to control the comments should be added to the Markdown or not (`<!-- Slide [n] ->`)

### Changed

- The parser now parses grouped elements (`<p:grpSp>`) recursively to find every base element inside of groups with `n` nested groups
- `SlideElements` are now sorted vertically before the Markdown conversion to preserve the visual hierachy
- `basic_usage.rs` now uses a second cmd parameter to control if images are extracted or not (for debug purposes)
---

## [0.3.0] - 2025-06-16

### Added

- Reworked the extraction of images by adding `ImageHandlingMode` to the `ParserConfig`. With this, users can decide to manually extract images and handle the logic [(#19)](https://github.com/nilskruthoff/pptx-parser/issues/19)
- New [example](https://github.com/nilskruthoff/pptx-parser/tree/master/examples) `manual_image_extraction.rs` to show how to handle images manually
- `ManualImage` struct to encapsulate data and meta data of images
- `ImageHandlingMode::Save` to save images in a given output path and adding context to the Markdown file [(#20)](https://github.com/nilskruthoff/pptx-parser/issues/20)

### Removed

- `image_extraction` from [examples](https://github.com/nilskruthoff/pptx-parser/tree/master/examples) directory (replaced by `manual_image_extraction.rs`)

### Changed

- Updated [README.md](https://github.com/nilskruthoff/pptx-parser/blob/master/README.md) to document new `ParserConfig` parameters

---

## [0.2.0] - 2025-06-15

### Added

- *multithreading* support for the parsing of slides [(#6)](https://github.com/nilskruthoff/pptx-parser/issues/6)
- `ParserConfig`: A config struct that increases the customizability for the devs [(#9)](https://github.com/nilskruthoff/pptx-parser/issues/9)
- Optional compression of extracted images [(#10)](https://github.com/nilskruthoff/pptx-parser/issues/10)
- Simple GitHub-Action to run all tests before merging a pull request [(`rust.yml`)](https://github.com/nilskruthoff/pptx-parser/blob/master/.github/workflows/rust.yml)
- unit tests for modules `parse_xml.rs`, `parse_rels.rs` and `slide.rs` [(#8)](https://github.com/nilskruthoff/pptx-parser/issues/8)
- `performance_test` example to run benchmarks 
- Started the Changelog [(#15)](https://github.com/nilskruthoff/pptx-parser/issues/15)

### Fixed

- minor bug fixes

### Changed

- Updated [README.md](https://github.com/nilskruthoff/pptx-parser/blob/master/README.md) to show the latest working examples and features
