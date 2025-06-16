# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [0.3.0] - _unreleased_

### Added

- Reworked the extraction of images by adding `ImageHandlingMode` to the `ParserConfig`. With this, users can decide to manually extract images and handle the logic [(#19)](https://github.com/nilskruthoff/pptx-parser/issues/19)
- New [example](https://github.com/nilskruthoff/pptx-parser/tree/master/examples) `manual_image_extraction.rs` to show how to handle images manually
- `ManualImage` struct to encapsulate data and meta data of images

### Removed

- `image_extraction` from [examples](https://github.com/nilskruthoff/pptx-parser/tree/master/examples) directory (replaced by `manual_image_extraction.rs`)

### Changed

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

- [README.md](https://github.com/nilskruthoff/pptx-parser/blob/master/README.md) updated to show the latest working examples and features