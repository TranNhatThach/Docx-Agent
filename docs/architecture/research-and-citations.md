# Research Assistant & Citation Provenance

## 1. Zero-Hallucination Policy
AI Agents must **never** fabricate paper titles, author names, DOIs, URLs, or publication years.
When evidence cannot be verified against a credible academic registry, the assistant explicitly reports `unsupported_warning`.

## 2. Supported Citation Standards
- **APA 7th Edition**: In-text `(Author, Year)` / Reference list format.
- **IEEE**: In-text `[1]` numerical citation / Numerical bibliography.
- **Academic-VN**: Vietnamese standard thesis citation format `(Tác giả, Năm)`.

## 3. Provenance Tracking
Every citation, generated image, and diagram block contains a `ProvenanceRecord`:
- `source_type`: `HUMAN`, `AGENT`, `IMPORT`, `RESEARCH`
- `creator`: Agent ID or username
- `source_url` / `doi`: Canonical persistent identifier
- `timestamp`: ISO-8601 creation timestamp
- `notes`: Audit explanation for why the artifact was selected
