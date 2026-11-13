# Parchi AI --- Project Overview, Problem, Solution, and Detailed Workflow

> **Purpose of this document:** Give Claude (or another coding
> assistant) a self-contained explanation of Parchi AI: the problem it
> solves, the intended solution, the end-to-end workflow, data rules,
> user experience, and implementation constraints.
>
> **Important for coding agents:** Treat this document as the product
> and workflow specification. Before changing code, inspect the existing
> repository and identify the current implementation. Preserve working
> functionality. Do not assume that this document describes the current
> code perfectly; where implementation and specification differ, explain
> the difference and propose the smallest safe change.

------------------------------------------------------------------------

## 1. Project summary

**Parchi AI** is a supplier-document information extraction and
organisation application.

Businesses receive information about a purchase or supply transaction
across multiple documents. These may include:

-   Supplier invoices
-   Purchase orders (POs)
-   Delivery challans
-   Goods Receipt Notes (GRNs)
-   Weighment slips
-   Lorry receipts (LRs) and transport documents
-   Supporting documents, attachments, and correspondence
-   Scanned PDFs, images, and spreadsheets

The information is often fragmented across documents, repeated,
inconsistently labelled, handwritten, or presented in different layouts.

Parchi AI processes a collection of documents (a *docket*), extracts all
identifiable information into a comprehensive **Main Metadata File**,
lets a client select the information they need, and exports that
information using the client's chosen template.

The central design principle is:

**Extract comprehensively first. Let the client decide what matters
afterwards.**

The extraction stage must not be limited by a predefined list of
business fields or by the client's current template.

------------------------------------------------------------------------

## 2. The problem Parchi AI solves

### 2.1 Information is scattered across documents

A single transaction may be represented by several documents. For
example:

-   The purchase order contains the ordered quantity and agreed rate.
-   The invoice contains the billed quantity, tax, and invoice total.
-   The delivery challan contains dispatch details.
-   The GRN contains received quantities.
-   The weighment slip contains gross, tare, and net weights.
-   The LR contains vehicle and transport information.

A person processing the docket must open each document, locate relevant
values, understand inconsistent labels, and transfer the information
into another format.

### 2.2 Document layouts and terminology vary

Different suppliers use different layouts and names for the same or
similar information. A supplier might use `Party Name`, another
`Supplier`, and another `Vendor Name`.

The same document type may also have different table structures,
headers, abbreviations, and ordering.

A rigid field list or fixed-position extraction process will miss
information when the layout changes.

### 2.3 Important information can be overlooked

Relevant information may appear in:

-   Headers and footers
-   Tables and individual line items
-   Stamps and seals
-   Handwritten annotations
-   Signatures and nearby text
-   Marginal notes
-   Small-print terms
-   Totals, subtotals, and tax breakdowns
-   Page continuations
-   Spreadsheet tabs, including hidden sheets when accessible

A process that extracts only familiar fields risks omitting details that
may become important later.

### 2.4 Repeated and conflicting values are meaningful

The same value may appear on multiple documents. A supplier name, order
number, date, or quantity may be repeated.

Two documents may also contain different values for what appears to be
the same concept. That difference may indicate a correction,
discrepancy, or legitimate distinction.

Parchi AI must not silently remove duplicates, select one conflicting
value, or assume that repeated values are unnecessary.

### 2.5 Every client may need a different output

One client may want invoice number, supplier name, tax amount, and net
weight. Another may need vehicle number, received quantity, purchase
order reference, and delivery date.

The application should not require a separate extraction pipeline for
each client. It should extract broadly once and let each client define
the output they need.

------------------------------------------------------------------------

## 3. The solution

Parchi AI separates the process into distinct stages:

1.  **Document intake:** Accept the complete set of source documents.
2.  **Exhaustive extraction:** Extract every identifiable information
    item from every accessible page and sheet.
3.  **Main Metadata creation:** Store each extracted item as an
    individual record, retaining repetitions.
4.  **Unique-label discovery:** Build a deduplicated list of labels for
    the user interface without changing the underlying records.
5.  **Client template selection:** Let the client create, select, or
    edit the output template.
6.  **Label matching:** Automatically match exact labels and show fuzzy
    suggestions for user approval.
7.  **Direct organisation and export:** Use approved mappings to place
    extracted values under the client's requested labels and export the
    result.
8.  **Validation and reporting:** Check that information has not been
    silently lost and report limitations or unresolved mappings.

The Main Metadata File is the comprehensive source dataset. The client
template is a view and export definition applied to that dataset.
Selecting a template must never modify or reduce the Main Metadata File.

------------------------------------------------------------------------

## 4. Core product principles

These principles apply throughout the application.

### 4.1 Exhaustiveness

Attempt to extract every identifiable information item from every
document page and every spreadsheet sheet that can be accessed.

Do not limit extraction to a predefined schema, a known set of invoice
fields, or the currently selected client template.

### 4.2 Preserve duplicates

Every occurrence is a separate extraction record, even if the same label
and value appear elsewhere.

Do not deduplicate the Main Metadata File.

A separate unique-label list may be created for display, but it must
never replace or modify the underlying records.

### 4.3 Do not guess

If a value is illegible, partially visible, ambiguous, or inaccessible,
do not invent a replacement.

Use an uncertainty-aware extraction result and an appropriate confidence
score. Make limitations visible to the user.

### 4.4 Confidence belongs to each extracted value

Every extracted record must have its own confidence score from 0 to 100.

The score represents confidence in that particular extracted value, not
confidence in the whole document or in a label mapping.

Where the interface displays confidence, use clear colour bands: -
**Green:** high confidence - **Yellow:** moderate confidence - **Red:**
low confidence

The exact thresholds should be consistent and documented in the
implementation. Do not imply that a colour makes an uncertain value
verified.

### 4.5 Keep extraction separate from organisation

Extraction identifies information in source documents.

Organisation determines which extracted records belong under the labels
requested by a client.

These are different operations and should remain separate in the
architecture.

### 4.6 Preserve provenance

Every extracted and exported value must retain the original source
document filename.

Do not discard this association when grouping values under client
template labels.

### 4.7 Avoid unnecessary AI calls

Use deterministic local code for tasks such as: - Exact string
matching - Fuzzy string similarity - Duplicate prevention in the UI -
Format conversion - Structural validation - Export generation

Use an AI model for document understanding and extraction where
required. Do not call an AI model merely to perform a task that can be
done reliably with local code.

------------------------------------------------------------------------

## 5. Main Metadata File

The Main Metadata File is the canonical, exhaustive extraction result.

### 5.1 Required record structure

Each extracted value must be represented by an individual record
containing exactly these four fields:

  -----------------------------------------------------------------------
  Field                               Meaning
  ----------------------------------- -----------------------------------
  `document_name`                     Original filename of the source
                                      document

  `label`                             Name assigned to the extracted
                                      information item

  `value`                             Extracted value, represented
                                      without inventing missing content

  `confidence_score`                  Confidence score from 0 to 100 for
                                      this value
  -----------------------------------------------------------------------

Example:

``` json
[
  {
    "document_name": "invoice_104.pdf",
    "label": "Invoice Number",
    "value": "INV-104",
    "confidence_score": 98
  },
  {
    "document_name": "purchase_order_104.pdf",
    "label": "Purchase Order Number",
    "value": "PO-104",
    "confidence_score": 99
  },
  {
    "document_name": "invoice_104.pdf",
    "label": "Supplier Name",
    "value": "ABC Industries",
    "confidence_score": 96
  },
  {
    "document_name": "delivery_challan_104.pdf",
    "label": "Supplier Name",
    "value": "ABC Industries",
    "confidence_score": 94
  }
]
```

<<<<<<< HEAD
This is illustrative. The actual extracted labels and values must come
from the documents.

### 5.2 Strict schema rule

Do not add page numbers, evidence excerpts, source coordinates, document
type, or other fields to the Main Metadata record unless the product
owner explicitly changes this specification.

The four fields above are the required Main Metadata record structure.

Internal processing may use additional temporary information if
necessary, but it must not leak into the four-field exported Main
Metadata schema.

### 5.3 Repeated records

The example deliberately includes two `Supplier Name` records. They must
both remain in Main Metadata because they originate from separate
document occurrences.

Do not use a dictionary keyed only by label. That would overwrite
repeated values.

Use a list/array of records or another structure that preserves multiple
occurrences.

### 5.4 Unique labels are a separate derived view

The application should derive a list of unique labels from all extracted
records.

For example, the records may contain `Supplier Name` ten times. The
label-selection interface should normally display that label once.

This is only a presentation and selection convenience. It must not
delete, merge, or alter the ten underlying records.

------------------------------------------------------------------------

## 6. Detailed end-to-end workflow

### Stage 1 --- Document intake

The user uploads a docket containing one or more supplier-related
documents.

Supported inputs should follow the application's existing file support.
Typical inputs include PDFs, images, and spreadsheets.

The application should: 1. Accept multiple files. 2. Retain each
original filename. 3. Validate file type and size. 4. Identify
unreadable or unsupported files. 5. Keep files associated with the
current processing session or docket. 6. Avoid silently ignoring a file
that failed validation.

If duplicate-upload detection exists, it may warn the user when
identical files are uploaded. It must not silently discard a file
without an explicit product decision.

### Stage 2 --- Exhaustive document processing and extraction

The selected vision/OCR/AI extraction provider processes the documents.

The extraction task is to identify all readable information, not merely
familiar fields.

The extraction instruction must cover:

-   Every accessible page in every PDF.
-   Every accessible image.
-   Every accessible spreadsheet sheet.
-   Hidden spreadsheet sheets where the file-reading method makes them
    accessible.
-   Printed text and handwriting.
-   Headers, footers, stamps, annotations, signatures, and marginal
    notes.
-   Tables, line items, quantities, rates, amounts, subtotals, totals,
    and tax details.
-   Repeated information and repeated occurrences.
-   Small text and footnotes.
-   Unusual labels, abbreviations, and supplier-specific terminology.

The model must not: - Restrict output to a fixed list of fields. -
Filter values according to a client template. - Remove repeated
information. - Resolve conflicting values by choosing one. - Fill in
unreadable values from assumptions. - Claim that an unprocessed page or
sheet was successfully read.

#### Processing completeness

A successful API response is not proof that the entire docket was
successfully processed.

The application should track processing status for each input and each
accessible page/sheet wherever the existing processing architecture
allows it.

Report: - Files successfully processed. - Files that failed. - Pages or
sheets that could not be accessed or processed. - Unreadable or
uncertain content, where detectable. - Partial results.

Do not report complete extraction when part of the docket was skipped.

### Stage 3 --- Main Metadata generation

Convert the extraction results into the canonical Main Metadata record
structure.

Each record contains: - `document_name` - `label` - `value` -
`confidence_score`

Preserve every record, including repeated labels and repeated values.

If a large docket must be processed in batches: 1. Process manageable
batches. 2. Retain the results from each batch. 3. Combine them without
overwriting records. 4. Validate the combined record count and
structure. 5. Report failed batches rather than silently returning a
partial result as complete.

The Main Metadata File must remain the authoritative extraction dataset
for later selection and export.

### Stage 4 --- Unique-label discovery

Read the Main Metadata records and derive the unique labels.

For example:

``` text
Supplier Name
Invoice Number
Invoice Date
Purchase Order Number
Vehicle Number
Gross Weight
Tare Weight
Net Weight
Tax Amount
Total Amount
```

The list is derived from actual extracted labels. It is not a hardcoded
list.

Rules: - Display each distinct label once. - Do not remove duplicate
records from Main Metadata. - Do not combine values at this stage. - Do
not infer that two different labels mean the same thing.

### Stage 5 --- Client template selection

The client can: - Create a new template. - Select an existing saved
template. - Edit a template. - Save a template for reuse.

A client template defines the output labels that the client wants.

Example:

``` json
[
  "Supplier",
  "Bill Number",
  "Bill Date",
  "PO Reference",
  "Net Weight",
  "Final Amount"
]
```

The exact spelling, capitalisation, spacing, and punctuation of these
labels are significant.

Do not silently rewrite a client's template labels.

### Stage 6 --- Label matching and selection

The application compares extracted labels against the selected client
template.

Matching has three categories.

#### A. Exact matches --- automatic

If an extracted label exactly matches a template label after the agreed
normalisation: - Ignore leading/trailing whitespace. - Compare
case-insensitively. - Automatically map the extracted label to the
existing template label. - Preserve the original template label exactly
in the final output.

Example:

``` text
Extracted label: " invoice number "
Template label:  "Invoice Number"
Result: automatic exact match
```

Normalisation is only for comparison. It must not alter the stored
source label or template label.

#### B. Similar labels --- suggestions only

Use RapidFuzz locally to compare labels that did not match exactly.

-   Show suggestions only when similarity is at least 70%.
-   Display the similarity score as a percentage.
-   Do not automatically accept fuzzy matches.
-   Let the user accept or reject each suggestion.
-   A similarity score is lexical similarity, not proof that two labels
    have the same business meaning.

Example:

``` text
Extracted label: "Vendor Name"
Template label:  "Supplier"
Similarity:      [calculated by RapidFuzz]
Action:          user decides whether to map
```

Do not assume that every similar-looking label is semantically
equivalent.

#### C. Unmatched labels

If no exact match or suggestion at or above 70% exists: - Display the
extracted label as unmatched. - Allow the user to add it to the client
template if that is what they want. - Do not silently discard it. - Do
not create a guessed mapping.

#### Interactive label pills

Extracted labels should be displayed as compact, clickable pills/chips.

Expected behaviour: - Clicking an available label adds it to the user's
template. - A label already added must be visually distinguishable. -
Prevent accidental duplicate additions. - Allow removal of a selected
label. - Update the interface immediately after an action. - Clearly
distinguish automatic exact matches, fuzzy suggestions, unmatched
labels, and labels already selected.

For fuzzy suggestions, make the proposed target template label and score
clear, and require the user to approve or reject the mapping.

**Important mapping distinction:** Adding a new extracted label to a
template and mapping an extracted label to an existing template label
are different actions. The UI must make the action clear. Never silently
replace an existing client template label with an extracted label.

Use RapidFuzz on the CPU. Do not introduce embeddings, a second AI call,
or a model download for this task.

### Stage 7 --- Build the client-specific result

Use the approved mappings and selected template labels to organise the
relevant Main Metadata records.

This stage should be deterministic and local where possible. The current
requested workflow does not require a **"Organize with Gemini"** button
or a Gemini organisation step.

For each client template label: 1. Find the extracted labels explicitly
mapped to it. 2. Retrieve all underlying Main Metadata records belonging
to those mapped labels. 3. Retain every occurrence. 4. Preserve the
original source document filename for each value. 5. Keep conflicting
values rather than choosing one. 6. Do not invent a value if no record
exists.

Example conceptual output:

``` json
{
  "Supplier": [
    {
      "document_name": "invoice_104.pdf",
      "value": "ABC Industries"
    },
    {
      "document_name": "delivery_challan_104.pdf",
      "value": "ABC Industries"
    }
  ],
  "Final Amount": [
    {
      "document_name": "invoice_104.pdf",
      "value": "118000.00"
    }
  ]
}
```

The example shows the organisation concept. The final file schema should
be implemented consistently across JSON, PSV, and XLSX and must preserve
the information listed above.

If a template label has no mapped source records, represent the absence
consistently (for example, an empty list in JSON or an empty value area
in a tabular export). Do not fabricate a value.

### Stage 8 --- Export

The export interface must contain exactly two primary panels.

#### Panel 1: Main Metadata

Contains: - A dropdown for the export format. - A simple export/download
action.

Format options: - JSON - PSV - XLSX

This panel exports the complete Main Metadata dataset.

It must include all records and all four required fields. It must not be
filtered by the client template.

#### Panel 2: Client Template

Contains: - A separate dropdown for the export format. - A simple
export/download action.

Format options: - JSON - PSV - XLSX

This panel exports the client-selected and mapped information under the
exact client template labels.

It must preserve: - Every applicable value occurrence. - Conflicting
values. - Source document filenames. - The client's exact template
labels.

#### Export behaviour

-   The two dropdowns are independent. Changing one must not change the
    other.
-   Do not create separate panels for JSON, PSV, and XLSX.
-   Do not require Gemini organisation to export.
-   Reuse export/conversion logic where practical.
-   Ensure that all three formats represent equivalent underlying
    information.
-   Do not modify the Main Metadata File when generating a client
    export.

**PSV means pipe-separated values**, using `|` as the delimiter.
Correctly escape or quote values containing the delimiter, quotation
marks, or line breaks so that records can be read back without
corruption.

### Stage 9 --- Validation and export integrity

Before offering a download, validate the generated output.

For Main Metadata: - Every record has exactly the four required
fields. - Confidence scores are within 0--100. - Repeated records have
not been overwritten. - The exported record count agrees with the source
record count. - Filenames, labels, and values have not been silently
changed.

For Client Template: - Every output label matches the client template
exactly. - Every selected and approved mapping is applied. - All source
records belonging to approved mappings are retained. - Duplicate and
conflicting values remain present. - Every value retains its source
document filename. - No unsupported or guessed mapping has been
introduced.

For all formats: - JSON is valid and parseable. - PSV is correctly
escaped and parseable. - XLSX contains the expected rows and columns. -
The same underlying information is represented in JSON, PSV, and XLSX.

If validation fails, show an actionable error. Do not present an invalid
or incomplete export as successful.

------------------------------------------------------------------------

## 7. User interface structure

The export area should be simple and contain two primary panels only:

``` text
EXPORT

+--------------------------------------+
| Main Metadata                        |
|                                      |
| Format: [ JSON  v ]                  |
|                                      |
|                       [ Export ]     |
+--------------------------------------+

+--------------------------------------+
| Client Template                      |
|                                      |
| Format: [ XLSX  v ]                  |
|                                      |
|                       [ Export ]     |
+--------------------------------------+
```

The selected formats above are examples, not defaults that must be
hardcoded.

The application should not display separate JSON, PSV, and XLSX export
panels.

The **"Organize with Gemini"** button must be removed from the
interface. Export should work using the selected template and approved
local mappings.

------------------------------------------------------------------------

## 8. Data integrity rules

These are non-negotiable.

  -----------------------------------------------------------------------
  Rule                                Required behaviour
  ----------------------------------- -----------------------------------
  Main Metadata schema                Exactly four fields per record

  Duplicate extraction records        Preserve all

  Duplicate values                    Preserve all occurrences

  Conflicting values                  Preserve all; do not choose
                                      silently

  Unique labels                       Deduplicate only in the derived UI
                                      label list

  Client template labels              Preserve exact text

  Exact matches                       Automatically map

  Fuzzy matches                       Suggest at 70% or higher; require
                                      user approval

  Unmatched labels                    Keep visible; do not discard
                                      silently

  Source filenames                    Retain with every
                                      extracted/exported value

  Missing information                 Do not invent

  Extraction confidence               Score each extracted value from 0
                                      to 100

  Main Metadata export                Complete, unfiltered dataset

  Client Template export              Only selected and approved mapped
                                      information

  Export formats                      JSON, PSV, XLSX

  Gemini organisation                 Not part of the requested export
                                      workflow
  -----------------------------------------------------------------------

------------------------------------------------------------------------

## 9. Suggested conceptual architecture

This is a logical separation of responsibilities, not a demand to
rebuild the current codebase.

``` text
                    USER UPLOADS DOCKET
                            |
                            v
                    DOCUMENT VALIDATION
                            |
                            v
                 DOCUMENT/PAGE/SHEET PROCESSING
                            |
                            v
                 EXHAUSTIVE OCR / AI EXTRACTION
                            |
                            v
                    MAIN METADATA RECORDS
                [document_name, label, value,
                     confidence_score]
                            |
                 +----------+----------+
                 |                     |
                 v                     v
          UNIQUE LABELS          COMPLETE METADATA
          FOR USER INTERFACE      FOR EXPORT
                 |
                 v
          CLIENT TEMPLATE
                 |
                 v
          LABEL MATCHING
           /     |      \
          /      |       \
       Exact   Fuzzy    Unmatched
       match  suggestion  labels
        |       |          |
        |    User review   User can
        |    accept/reject add label
        +-------+----------+
                |
                v
        APPROVED MAPPINGS
                |
                v
       CLIENT-SPECIFIC SELECTION
                |
                v
        CLIENT TEMPLATE DATA
                |
                v
       VALIDATION AND EXPORT
                |
          +-----+-----+
          |           |
          v           v
     MAIN METADATA  CLIENT TEMPLATE
       EXPORT          EXPORT
      JSON/PSV/XLSX   JSON/PSV/XLSX
```

Possible module responsibilities, adapted to the existing repository:

-   **Upload/input layer:** file validation and docket/session handling.
-   **Document processing layer:** reading PDFs, images, and
    spreadsheets.
-   **Extraction layer:** OCR/vision model calls and extraction-result
    parsing.
-   **Metadata layer:** canonical four-field records and completeness
    checks.
-   **Label layer:** unique-label derivation and label display.
-   **Template layer:** create, edit, save, and load client templates.
-   **Matching layer:** exact matching, RapidFuzz suggestions, and
    approved mappings.
-   **Organisation layer:** deterministic grouping of records under
    template labels.
-   **Export layer:** JSON, PSV, and XLSX generation.
-   **Validation layer:** schema, count, mapping, and export-integrity
    checks.

Do not create new modules merely to match these names if the existing
project already has a sensible structure. Prefer small changes within
the current architecture.

------------------------------------------------------------------------

## 10. Important implementation constraints for coding assistants

When modifying Parchi AI:

1.  Inspect the repository, entry points, dependencies, and current
    workflow before editing.
2.  Identify which features already work. Preserve them.
3.  Do not rewrite the application from scratch.
4.  Do not replace the existing framework or architecture without
    explicit approval.
5.  Do not alter the extraction process when implementing UI-only export
    changes.
6.  Do not change the Main Metadata record schema.
7.  Do not deduplicate the underlying extraction records.
8.  Do not make fuzzy matches automatic.
9.  Do not use an AI call for local label similarity or file conversion.
10. Do not reintroduce the "Organize with Gemini" button.
11. Do not create separate export panels for each format.
12. Do not make the two format dropdowns share state.
13. Do not silently swallow errors or report partial extraction as
    complete.
14. Do not log API keys, document contents, or sensitive extracted
    values unnecessarily.
15. Make incremental changes and test each affected workflow.
16. Report which files were changed, what changed, and what tests were
    run.
17. If the current code conflicts with this specification, explain the
    conflict before making a risky change.

### API key handling

The product requirement is that the user should be able to enter the API
key through the application's UI and have it securely persist across
restarts.

Do not require the user to enter the key in a terminal, source file,
`.env` file, or environment variable.

Use an appropriate operating-system credential store where supported. Do
not silently fall back to storing the key in plaintext. Never expose the
key in logs, exports, or the visible interface after it has been saved.

------------------------------------------------------------------------

## 11. Testing checklist

### Extraction and metadata

-   [ ] Multiple documents can be uploaded in one docket.
-   [ ] Every accessible page is processed.
-   [ ] Every accessible spreadsheet sheet is processed.
-   [ ] Handwritten and printed content are both attempted.
-   [ ] Repeated values remain separate records.
-   [ ] Conflicting values remain separate records.
-   [ ] Each record has exactly four required fields.
-   [ ] Every confidence score is between 0 and 100.
-   [ ] Failed or unreadable content is reported honestly.
-   [ ] Large/batched processing does not overwrite earlier results.

### Labels and templates

-   [ ] Unique labels are derived from actual extracted labels.
-   [ ] Main Metadata is not changed when labels are deduplicated for
    display.
-   [ ] New templates can be created and saved.
-   [ ] Existing templates can be selected and edited.
-   [ ] Exact matches are automatically mapped.
-   [ ] Case and surrounding whitespace normalisation works for
    comparison.
-   [ ] Fuzzy suggestions below 70% are hidden.
-   [ ] Fuzzy suggestions at or above 70% show their score.
-   [ ] Fuzzy suggestions require user approval.
-   [ ] Unmatched labels remain visible.
-   [ ] Clicking a pill adds the intended label.
-   [ ] Duplicate additions are prevented.
-   [ ] Selected labels can be removed.
-   [ ] Template labels retain exact text.

### Client organisation

-   [ ] Every approved mapping retrieves all applicable records.
-   [ ] Repeated values remain present.
-   [ ] Conflicting values remain present.
-   [ ] Source filenames remain associated with values.
-   [ ] Missing values are not fabricated.
-   [ ] No unapproved fuzzy mapping is applied.

### Export

-   [ ] Main Metadata panel has JSON, PSV, and XLSX options.
-   [ ] Client Template panel has JSON, PSV, and XLSX options.
-   [ ] There are only two primary export panels.
-   [ ] Each panel has its own independent format selection.
-   [ ] Main Metadata export contains all records.
-   [ ] Client Template export contains selected/mapped information
    only.
-   [ ] JSON output is valid.
-   [ ] PSV escaping handles pipes, quotes, and newlines.
-   [ ] XLSX output preserves rows and values.
-   [ ] All formats represent equivalent information.
-   [ ] Export does not require Gemini organisation.
-   [ ] The "Organize with Gemini" button is absent.
-   [ ] Exporting does not mutate the Main Metadata File.

------------------------------------------------------------------------

## 12. Definition of success

Parchi AI succeeds when a user can upload a mixed supplier-document
docket, obtain a comprehensive and traceable set of extracted records,
select the labels relevant to their needs, approve any ambiguous label
mappings, and export both the complete Main Metadata and the
client-specific template in JSON, PSV, or XLSX.

The user should not need to repeat extraction just because a different
client wants a different output.

**In one sentence: Parchi AI turns fragmented supplier documents into an
exhaustive, duplicate-preserving, confidence-scored metadata dataset
that can be selectively mapped and exported into different client
templates.**
=======
Set `GEMINI_API_KEY` or add API keys in the application settings to enable Gemini extraction. Uploads are processed one document at a time; PDFs are chunked by page, and XLSX workbooks include hidden worksheets when they are readable. Local runtime databases and uploaded documents are intentionally ignored by Git.
>>>>>>> refs/remotes/origin/main
