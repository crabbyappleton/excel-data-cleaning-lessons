# Excel & Data-Cleaning Lessons: Syllabus for Lessons 02–06

This five-lesson sequence follows **Lesson 01: Structured Data Starts Here**. It is designed for a complete beginner using Excel for the web, with an emphasis on practical work in data collection, cleaning, and annotation. Each lesson builds on the same core rule introduced in Lesson 01:

> **One row = one record. One column = one attribute.**

By the end of Lesson 06, you will be able to take a small, imperfect dataset, organize it as an Excel Table, identify common quality problems, clean text and duplicate records, standardize annotation labels, and run simple quality checks before sharing the data.

| Lesson | Title | Main Excel skills | Practical outcome |
|---|---|---|---|
| 02 | Build a Table You Can Work With | Excel Tables, filters, sorting, freeze panes | A sortable, filterable customer-comment table |
| 03 | Find What Needs Cleaning | Blanks, duplicates, inconsistent formats, filter views | A documented list of data-quality issues |
| 04 | Clean Text, Dates, and IDs | Find and Replace, text functions, date formatting, IDs | A standardized working dataset |
| 05 | Label Data Consistently | Data validation, drop-down lists, conditional formatting | An annotation sheet with controlled labels |
| 06 | Check Your Work Before You Share It | `COUNTIF`, `IF`, filters, simple summaries | A basic quality-control checklist and review summary |

## Lesson 02 — Build a Table You Can Work With

The second lesson converts a plain range of cells into an **Excel Table**. A table is the working container for the rest of the series: it keeps headings connected to records, adds filter controls, and expands as new rows are added. The lesson introduces filtering as a way to inspect a subset of records and sorting as a way to put records in a useful order.

| Component | Plan |
|---|---|
| Learning objective | Turn a small range into a named Excel Table and use its filter arrows confidently. |
| Core actions | Select the data range; choose **Insert → Table**; confirm that the range has headers; sort Record ID; filter by one Topic value; freeze the top row. |
| Practice data | Six customer comments with Record ID, Comment, Topic, Sentiment, and Reviewed? columns. |
| Mini challenge | Filter to show only records whose `Reviewed?` value is `No`, then sort those records by Topic. |
| Finished product | A table called `CommentsTable` with a clean header row and usable filters. |

The key habit is to avoid placing notes, totals, decorative titles, or blank rows inside the table. Keep explanatory notes outside the table so that filters and sorts act only on the records.

## Lesson 03 — Find What Needs Cleaning

Lesson 03 focuses on diagnosis before editing. You will learn to recognize the most common problems in beginner datasets: blank required cells, duplicate IDs, inconsistent category labels, and mixed date formats. Rather than changing everything at once, you will use filters and a short issue log to identify what requires attention.

| Component | Plan |
|---|---|
| Learning objective | Inspect a dataset systematically and identify specific, fixable quality problems. |
| Core actions | Filter blanks in a column; sort to group similar values; use conditional formatting to highlight duplicate IDs; create an `Issue Notes` column. |
| Practice data | A 15-row contact or customer-feedback dataset containing intentional blanks, duplicated Record IDs, and variants such as `delivery`, `Delivery`, and `DELIVERY`. |
| Mini challenge | Find all duplicate IDs and all records missing a Topic label without deleting or overwriting anything. |
| Finished product | An issue log that states the record, the problem, and the next action. |

Use this three-question inspection routine: **Is every required value present? Is each value in the expected format? Is each identifier unique?** That routine will remain useful in every later lesson.

## Lesson 04 — Clean Text, Dates, and IDs

Lesson 04 turns diagnosis into careful repair. You will clean common text inconsistencies, standardize dates, and create reliable identifiers. The goal is not to erase useful detail; it is to make values comparable so Excel can sort, filter, and count them correctly.

| Component | Plan |
|---|---|
| Learning objective | Standardize common text and date problems while preserving the original records. |
| Core actions | Use **Find and Replace** for deliberate global corrections; apply consistent date formatting; use `TRIM` to remove extra spaces; use `UPPER`, `LOWER`, or `PROPER` when appropriate; format IDs as text when leading zeroes matter. |
| Practice data | A small table with leading/trailing spaces, mixed capitalization, three date styles, and Record IDs such as `1`, `002`, and `0003`. |
| Mini challenge | Create a clean Topic column from inconsistent topic entries and ensure all Record IDs display as three digits. |
| Finished product | A cleaned copy of the dataset plus a short record of the transformations used. |

A useful safety rule is to keep a raw column when a transformation may change meaning. For example, retain `Comment (Original)` and create `Comment (Clean)` if you are removing extra spaces or correcting obvious formatting problems.

## Lesson 05 — Label Data Consistently

Lesson 05 introduces **data annotation**: assigning controlled labels to records. You will build drop-down lists so labels stay consistent, apply conditional formatting so unfinished work is easy to spot, and use a review status so annotation work can be tracked without confusion.

| Component | Plan |
|---|---|
| Learning objective | Design and use consistent annotation labels for categories, sentiment, and review status. |
| Core actions | Create a small label list; use **Data Validation** to add drop-downs; apply conditional formatting to blanks or `Needs Review`; add a simple annotation guide above or beside the table. |
| Practice data | Ten customer comments that need a Topic, Sentiment, and Reviewed? label. |
| Mini challenge | Label all ten comments using only the approved values, then filter for records needing review. |
| Finished product | An annotation-ready table with controlled drop-down values and visible review cues. |

Start with a short, mutually exclusive label set. For example, the Topic column may use **Delivery**, **Account**, **Product**, and **Billing**. If a record truly does not fit, use **Other** and add a note rather than creating a new category impulsively.

## Lesson 06 — Check Your Work Before You Share It

Lesson 06 introduces lightweight quality control. You will use formulas and filters to check whether annotation is complete, whether IDs repeat, and how many records fall into each category. This is the first step from “I entered data” to “I can explain why this data is ready to use.”

| Component | Plan |
|---|---|
| Learning objective | Build simple checks that reveal incomplete, duplicate, or unusual records before the dataset is shared. |
| Core actions | Use `COUNTIF` to count labels or duplicate IDs; use `IF` to flag missing required fields; filter flagged rows; create a small category summary. |
| Practice data | The annotated customer-comment table from Lesson 05. |
| Mini challenge | Add a `Quality Check` column that returns `Ready` only when Record ID, Comment, Topic, Sentiment, and Reviewed? are all present. |
| Finished product | A cleaned and annotated dataset with a simple readiness check and label counts. |

The central formula pattern for this lesson is:

```excel
=IF(AND(A2<>"",B2<>"",C2<>"",D2<>"",E2<>""),"Ready","Needs Review")
```

This formula checks whether five required cells contain a value. You will adapt it to match the columns in your own table, then filter the `Needs Review` records until the table is ready to hand off.

## Suggested Pace and Project Thread

Plan for approximately **45–60 minutes per lesson**, including practice. Each lesson should use the same customer-comment dataset so that you see how a dataset improves through a realistic workflow instead of starting over each time.

| After completing | You should be able to say |
|---|---|
| Lesson 02 | “I can create a table and inspect records with filters and sorting.” |
| Lesson 03 | “I can spot missing, duplicate, and inconsistent values before I change them.” |
| Lesson 04 | “I can make text, dates, and IDs consistent without losing the original information.” |
| Lesson 05 | “I can create controlled labels for annotation and track work that needs review.” |
| Lesson 06 | “I can run basic checks to decide whether my dataset is ready to share or analyze.” |

After Lesson 06, the natural next unit is **collecting and importing data**: bringing a CSV into Excel, checking column types, and preparing it for the same cleaning and annotation workflow.
