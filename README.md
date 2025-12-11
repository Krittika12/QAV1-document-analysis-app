# Automated Quality Report Verification Tool

### Streamlining Multi-Sheet Excel Validation with NLP + Workflow Automation

## 📌 Problem Statement

Honda receives a **high volume of complex quality analysis reports** from over **50+ parts suppliers**.
Each supplier provides a **25-worksheet Excel report**, where every worksheet includes **dozens to hundreds of technical quality-check questions**.

Manual verification required:

* **30–60 minutes per file**
* High cognitive load and repetitive work
* Risk of **human errors**, inconsistent judgment, and delayed approval cycles

These delays directly affected:

* Supplier trust
* Manufacturing timelines
* Overall operational efficiency

There was an urgent need for an automated system that could:

* Parse and validate **25-sheet Excel files**
* Check hundreds of parameters with high accuracy
* Reduce review time by **80–90%**
* Ensure consistency, traceability, and structured decision-making

---

## 💡 Initial Vision

The idea was to build an **AI-assisted verification layer** on top of the existing supplier report format.

The approach used:

* A database of historically validated answers
* Question-wise keyword profiles representing correct responses
* NLP techniques for **keyword + semantic similarity analysis**

For every new report, the system would:

* Extract answers across 25 worksheets
* Run NLP-based **semantic similarity checks**
* Identify compliant vs. non-compliant responses
* Produce **visual flags** for quick reviewer attention

This shifted the process from a fully manual workflow to a **data-driven, semi-automated quality assurance pipeline**.

---

## 🛠️ Solution Overview
<img width="886" height="493" alt="image" src="https://github.com/user-attachments/assets/1c43c8dd-ebd7-449d-8912-4fd175958ec9" />
I built a full **desktop workflow automation tool** that performs end-to-end quality-report verification.

### ✔ Key Features

#### **1. Automated Excel Processing**

* Upload multi-sheet Excel files
* NLP-powered keyword + semantic similarity validation
* Automated identification of incorrect/incomplete answers
* Output: a **color-coded reviewer-ready Excel file**

#### **2. Dynamic Keyword Database (CRUD Enabled)**

To handle evolving supplier documents, I implemented a flexible keyword management system allowing engineers to:

* Insert a new question
* Delete a question
* Delete multiple questions
* Update worksheet/row mappings
* Update question text
* Update keywords
* Insert or delete keywords

This ensured:

* No hardcoded logic
* No developer intervention needed for changes
* Future-proofing as supplier report formats evolve

## Database Structure
<img width="842" height="352" alt="image" src="https://github.com/user-attachments/assets/12945078-606d-4abf-a6a8-2c83ba98d014" />

The database schema is designed to keep the system fully adaptable to changes in supplier documentation. Since the order of questions within each worksheet may change frequently, the structure supports fast insertion, deletion, and reordering without rewriting the entire worksheet.

### ✔ Doubly Linked List for Questions
- Each question record maintains two references:
  - `prev_q_id` — pointer to the previous question  
  - `next_q_id` — pointer to the next question
- Enables O(1) insertion and deletion by updating surrounding links.
- Supports efficient reordering without shifting row numbers or rebuilding worksheets.

### ✔ Head Question Reference
- Each worksheet stores a `head_question_id`, pointing to the first question in its linked list.
- Serves as the entry point for ordered traversal.

### ✔ Unique Correct Keywords per Question
- Every question has a list of unique correct keywords stored in the `keywords` table.
- A composite uniqueness constraint (`question_id`, `keyword_text`) prevents duplicates.
- Cascade delete ensures keywords are removed automatically when a question is deleted.

### ✔ Benefits of This Design
- Fully dynamic question ordering  
- Fast insertion and deletion  
- No row renumbering required  
- Clean relational structure with predictable traversal  
- Data integrity during frequent updates  
- Scales easily as worksheets or workbooks evolve

## Processing Excel Files: NLP

The extracted answers from supplier reports contained significant textual noise. To ensure reliable keyword extraction, data profiling and regex-based exploratory analysis were performed to identify recurring noise patterns. These patterns were grouped into 10 categories, covering formatting artifacts, document codes, numbering systems, and irrelevant metadata.

A regex-driven cleaning pipeline was built to normalize text before keyword extraction. After cleaning, only meaningful keywords were inserted into the database, improving accuracy when matching supplier responses to expected answers.

### ✔ Identified Noise Categories (10 Types)

1. **Numbered keyword artifacts**  
   Removal of keywords prefixed with numbering formats (e.g., `1.`, `1)`, `①`, etc.).

2. **規定 / 規程 / 帳票 / 手順 / 該当項 + digit**  
   Patterns like `規定1`, `手順3`, etc., excluded because they reference document sections rather than answer content.

3. **d.~~~~ patterns**  
   Detection of phrases beginning with a digit followed by a period (e.g., `1.手順説明`, `2.工程管理`).

4. **Date patterns**  
   Filtering of common date formats (`YYYY/MM/DD`, `YY-MM-DD`, etc.).

5. **Document code pattern: `XXX-XXX-XXXXX-XXXX`**  
   Removal of long structured alphanumeric codes classified as non-semantic noise.

6. **Codes wrapped in parentheses**  
   Examples like `(AB-12345)` or `(XX-9999)` removed to avoid misidentification as keywords.

7. **Japanese document counters**  
   Units such as `件`, `号`, `章`, `項` filtered when used as document references.

8. **Other digit-based codes**  
   Exclusion of short numeric or mixed codes (e.g., `101`, `12A3`).

9. **Loose bracket noise**  
   Removal of unpaired or empty brackets such as `()`, `[]`, `<>`.

10. **`.xlsx` file references**  
   Cleaning of filenames or spreadsheet references irrelevant to semantic meaning.
## Handling Semantically Similar Keywords (Embedding-Based Matching)

To handle cases where suppliers use different words with similar meaning, a semantic similarity component was integrated using the **cl-tohoku/bert-base-japanese** model. This enabled recognition of conceptually equivalent keywords instead of relying solely on exact string matching.

### ✔ Embedding Generation
- Dense vector embeddings were generated for each keyword using BERT-based Japanese embeddings.
- These embeddings capture semantic meaning, enabling flexible comparison between supplier answers and expected keywords.

### ✔ Semantic Similarity (Cosine Similarity)
- A cosine similarity matrix was computed between extracted keywords and expected keywords in the database.
- Keywords exceeding a similarity threshold were considered valid matches even when the wording differed.
  - Example: `作業内容` ↔ `作業手順`  
  - Example: `点検` ↔ `チェック`  
  - Includes synonyms, verb variations, and kanji/kana differences.
- This removed the need for manual synonym lists and improved linguistic robustness.

## Caching Embeddings for Performance Optimization

Initially, embeddings were recomputed for every keyword, resulting in ~30 minutes of processing time per Excel file. To optimize performance, an embedding cache was introduced.

### ✔ Dictionary-Based Embedding Cache
- Each computed embedding was stored in a Python dictionary:
  - `cached_embeddings[keyword] = embedding_vector`
- When a keyword appeared again, the system reused the cached embedding instead of recomputing it.

### ✔ Result
- Processing time reduced from **~30 minutes → 2–3 minutes** per Excel file.
- Improved scalability for large batches and repeated processing workflows.

## Final Impact

The automation system transformed a previously manual, error-prone verification process into a fast, reliable, and scalable workflow. By combining a dynamic database design with NLP-driven text cleaning, semantic embedding matching, and performance-optimized processing, the solution delivered measurable operational improvements:

- **90% reduction in verification time** per supplier report (from 60 minutes to under 5 minutes).  
- **High accuracy** in keyword detection through regex-based noise removal and semantic similarity matching.  
- **Consistent and reliable validation**, minimizing human errors and improving supplier communication.  
- **Fully adaptable structure** that supports changing worksheet formats without rewriting logic.  
- **Scalable pipeline** capable of handling large batches of 25-worksheet reports efficiently.  
- **Reduced workload and faster approval cycles**, directly improving trust and collaboration with suppliers.  
- **Future-proof architecture**, enabling continuous updates to questions, keywords, and formats with minimal effort.

Overall, the project significantly improved the quality analysis workflow, strengthened supplier relationships, and established a robust foundation for future intelligent document-processing initiatives.

## System Screenshots
<img width="670" height="390" alt="image" src="https://github.com/user-attachments/assets/f07098a6-501c-4c27-99b0-f5d3baebd1a5" />

<img width="895" height="709" alt="image" src="https://github.com/user-attachments/assets/b99fa72c-d380-4dfc-856a-4403f394c1cd" />

<img width="811" height="650" alt="image" src="https://github.com/user-attachments/assets/dc87f4d7-f108-49ea-97f1-f84139ca7b98" />

<img width="885" height="522" alt="image" src="https://github.com/user-attachments/assets/3aa11f21-df17-46d3-811f-8163a05a2531" />

<img width="887" height="518" alt="image" src="https://github.com/user-attachments/assets/a56d80ad-3923-4bd2-8f4c-d56c00fa1d5e" />

Color coded processed file
<img width="1826" height="1092" alt="image" src="https://github.com/user-attachments/assets/419a9960-3536-4490-aa9e-2618464f252b" />
