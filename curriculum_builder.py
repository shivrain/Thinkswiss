#!/usr/bin/env python3
"""
Comprehensive 2-hour/day curriculum builder for:
0. CS & Programming Foundations (NEW)
1. Python  2. R  3. QGIS & Spatial Thinking
4. Data Analysis & Statistical Modelling  5. Advanced Econometrics
+ Spatial Python & Spatial R integrated throughout
+ Interpretation & Analytical Thinking module (NEW)
"""

import openpyxl
from openpyxl.styles import (
    Font, PatternFill, Alignment, Border, Side, GradientFill
)
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.styles.numbers import FORMAT_TEXT

# ─── colour palette ────────────────────────────────────────────────────────────
CLR = {
    "header_bg":   "1F3864",   # dark navy
    "header_font": "FFFFFF",
    "alt_row":     "EBF3FB",   # light blue tint
    "white":       "FFFFFF",
    "accent1":     "2E75B6",   # medium blue (Python)
    "accent2":     "375623",   # dark green  (R)
    "accent3":     "7030A0",   # purple      (QGIS)
    "accent4":     "843C0C",   # brown       (Stats)
    "accent5":     "C00000",   # dark red    (Econometrics)
    "accent6":     "00B0F0",   # cyan        (Spatial)
    "tab_py":      "4472C4",
    "tab_r":       "70AD47",
    "tab_qgis":    "9E480E",
    "tab_stats":   "ED7D31",
    "tab_econ":    "FF0000",
    "tab_spy":     "00B0F0",
    "tab_sr":      "375623",
    "tab_over":    "1F3864",
}

THIN = Side(style='thin', color="BFBFBF")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)


def hdr_fill(hex_col):
    return PatternFill("solid", fgColor=hex_col)


def alt_fill():
    return PatternFill("solid", fgColor=CLR["alt_row"])


def white_fill():
    return PatternFill("solid", fgColor=CLR["white"])


def hdr_font(bold=True):
    return Font(name="Calibri", bold=bold, color=CLR["header_font"], size=11)


def body_font(bold=False, size=10):
    return Font(name="Calibri", bold=bold, color="000000", size=size)


def wrap_align(horizontal="left", vertical="top"):
    return Alignment(wrap_text=True, horizontal=horizontal, vertical=vertical)


def write_row(ws, row_num, values, is_header=False, is_alt=False, col_start=1):
    for col_offset, val in enumerate(values):
        cell = ws.cell(row=row_num, column=col_start + col_offset, value=val)
        if is_header:
            cell.fill = hdr_fill(CLR["header_bg"])
            cell.font = hdr_font()
        elif is_alt:
            cell.fill = alt_fill()
            cell.font = body_font()
        else:
            cell.fill = white_fill()
            cell.font = body_font()
        cell.alignment = wrap_align()
        cell.border = BORDER


# ══════════════════════════════════════════════════════════════════════════════
#  CURRICULUM DATA
# ══════════════════════════════════════════════════════════════════════════════

# Each entry: (day, week, module, topic, subtopics, what_to_cover,
#              primary_resource, secondary_resources, dataset_practice,
#              assignment, paper_to_read, key_concepts)

CS_DATA = [
    # ── MODULE 1: HOW COMPUTERS WORK ──────────────────────────────────────────
    (1, 1, "M1: How Computers Work", "Hardware, Memory, and What a Program Really Is",
     "CPU (processor); RAM vs storage (SSD/HDD); binary representation of numbers and text; what the OS does; processes and threads (intuition); why Python is slow vs C (interpreted vs compiled); what happens when you run a script",
     "You will write a lot of code before you understand what it is actually doing at the machine level. This day fills that gap. Key insight: a program is a sequence of instructions given to the CPU. RAM is a temporary scratchpad — when you load a CSV into Python, it goes into RAM. If your dataset is larger than RAM, things break. Understanding this explains why rasters crash your computer but a CSV does not. Knowing binary representation explains why R has integers vs doubles, and why floating point arithmetic produces surprises (0.1 + 0.2 ≠ 0.3 exactly).",
     "Crash Course Computer Science (YouTube, Ep.1-4) | https://youtube.com/playlist?list=PL8dPuuaLjXtNlUrzyH5r6jN9ulIgZBpdo",
     "CS50 Week 0 (Harvard, free) | https://cs50.harvard.edu/x/2024/weeks/0/ | How Computers Work — Khan Academy | https://www.khanacademy.org/computing/computers-and-internet/xcae6f4a7ff015e7d:computers | But How Do It Know? (book, Ch.1-3 only) | https://www.amazon.com/But-How-Know-Principles-Computers/dp/0615303765",
     "No dataset. Practice: open Python/R and type 0.1 + 0.2. Then type 2**53 + 1 == 2**53 in Python. Then open Task Manager / Activity Monitor and watch RAM usage while loading a large CSV. These are empirical experiments about computer behaviour.",
     "Answer these in writing (look up anything you don't know): (1) What is the difference between RAM and a hard drive? (2) Why does opening a 4GB raster file sometimes crash your laptop if it has 8GB RAM? (3) Why is Python slower than C at basic arithmetic? (4) What does 'running a script' actually mean step by step? (5) What does your OS do that your Python code doesn't have to?",
     "None — this is the day you read the hardware explainers. CS50 Week 0 lecture is 1.5 hours; watch it at 1.5× speed.",
     "CPU, RAM, storage, binary, OS, process, interpreted vs compiled, floating point"),

    (2, 1, "M1: How Computers Work", "The Command Line: Your Research Superpower",
     "Terminal / command prompt basics; file system navigation (pwd, ls/dir, cd); creating and moving files (mkdir, cp, mv, rm); running scripts from terminal (python3 script.py, Rscript script.R); file paths (absolute vs relative); environment variables (PATH); piping (|) and redirection (>); bash scripts (basics only)",
     "The command line is how researchers interact with computing environments that have no graphical interface — remote servers, HPC clusters, cloud VMs like the one you are reading this on. It is also faster for repetitive operations: renaming 200 data files takes 3 seconds on the command line and 20 minutes by hand. For your thesis: downloading datasets, running scripts, and organising project directories are all more efficient from the terminal. You do not need to master it — you need to not be afraid of it.",
     "The Unix Shell (Software Carpentry, free) | https://swcarpentry.github.io/shell-novice/ — work through Episodes 1-5",
     "Command Line Crash Course (Zed Shaw, free) | https://learncodethehardway.org/unix/ | Bash Tutorial | https://linuxcommand.org/lc3_learning_the_shell.php | Missing Semester of CS Education (MIT, free) | https://missing.csail.mit.edu/2020/course-shell/",
     "No external dataset. Use your own computer's file system as the dataset. Navigate to your thesis folder (or create one), organise it using only the command line.",
     "Complete these command line tasks: (1) Navigate to your home directory and list all files including hidden ones; (2) Create a directory structure: thesis/data/raw, thesis/data/processed, thesis/code, thesis/output/figures using a single mkdir -p command; (3) Download a file from the internet using curl or wget (e.g., a CSV from a URL); (4) Count the number of lines in a CSV file using wc -l; (5) Find all .csv files in a directory and its subdirectories using find; (6) Write a one-line bash script that creates a dated backup of a file.",
     "Read: MIT Missing Semester — Course Shell lecture notes | https://missing.csail.mit.edu/2020/course-shell/ (30 min read).",
     "pwd, ls, cd, mkdir, cp, mv, rm, absolute/relative path, environment variable, pipe, redirection"),

    (3, 2, "M2: Programming Concepts", "How Programming Languages Work — Across Python, R, and STATA",
     "Interpreted vs compiled languages; dynamic vs static typing; object-oriented vs functional vs procedural paradigms; why Python and R are different design philosophies; what a library/package is and how it is loaded; namespaces and scope; how Python imports work; what a REPL is; scripting vs interactive computing",
     "Python was designed as a general-purpose language that is readable and pragmatic. R was designed by statisticians as a domain-specific language for statistical computing — its defaults assume you are doing data analysis. STATA was designed for social scientists and economists with a command-based interface. Understanding these design philosophies explains why R has vectors as the base object (not scalars), why Python needs pandas for data tables, and why STATA does one dataset at a time. Knowing this helps you choose tools wisely and understand error messages that reference these concepts.",
     "Python vs R vs STATA — Towards Data Science | https://towardsdatascience.com/python-vs-r-for-data-science-6a83e4541571 | Python Data Model (official) | https://docs.python.org/3/reference/datamodel.html",
     "Programming Language Pragmatics Ch.1 (library) | Comparison of R and Python for data science | https://www.datacamp.com/blog/python-vs-r-for-data-science | Why R is hard to learn | https://r4stats.com/articles/why-r-is-hard-to-learn/",
     "No dataset. Experiment: write the same operation in Python, R, and if possible STATA — compute mean of a list of 10 numbers. Notice: how many lines? What does the syntax require?",
     "Side-by-side comparison exercise: for each of the following tasks, write the solution in both Python AND R. (1) Create a list/vector of numbers 1-10; (2) Compute mean, median, sd; (3) Filter to keep only values > 5; (4) Write a function that takes a list and returns its z-scores; (5) Load a CSV and print the first 5 rows. Write a 200-word reflection: what feels more natural in each language and why?",
     "Read: Gentleman, R. & Ihaka, R. (1996). 'R: A Language for Data Analysis and Graphics.' Journal of Computational and Graphical Statistics, 5(3), 299-314. — the original R paper. It is short and explains the design decisions.",
     "interpreted, compiled, dynamic typing, paradigms, namespace, scope, REPL, package/library loading"),

    (4, 2, "M2: Programming Concepts", "Algorithms and Computational Thinking",
     "What is an algorithm; pseudocode; Big-O notation (O(1), O(n), O(n²)) — intuition only, no maths; why vectorisation beats loops; sorting algorithms (concept only); recursion (concept + one example); searching; why your for loop over 1 million rows is slow; divide-and-conquer thinking",
     "Computational thinking is the skill of decomposing a problem into steps a computer can execute. Big-O tells you whether your analysis will run in 1 second or 1 day on large data — O(n²) means if you double your data size, runtime quadruples. For social science data: loops inside loops on large datasets are the most common performance mistake. Understanding why vectorised NumPy/R operations are fast (they run in compiled C under the hood) lets you write better code without being a computer scientist.",
     "Crash Course Computer Science Ep.13 (Algorithms) | https://youtu.be/rL8X2mlNHPM | Khan Academy: Algorithms | https://www.khanacademy.org/computing/computer-science/algorithms",
     "CS50 Week 3 (Algorithms, Harvard, free) | https://cs50.harvard.edu/x/2024/weeks/3/ | Big-O cheat sheet | https://www.bigocheatsheet.com/ | Computational Thinking for Everyone | https://www.cs.cmu.edu/~wing/publications/Wing06.pdf",
     "Empirical experiment: time a for loop vs vectorised operation in Python/R on a dataset of 1M rows. Use time.time() or system.time().",
     "Timing experiment: (1) Create a list of 500,000 random numbers in Python; (2) Compute the sum using (a) a for loop, (b) sum(), (c) np.sum(); (3) Record time for each using time.time(); (4) Create a data frame in R with 500,000 rows; (5) Compute row-wise mean using (a) a for loop, (b) rowMeans(); (6) Record time; (7) Write a paragraph explaining WHY the vectorised versions are faster, referencing what you learned about how computers work.",
     "None",
     "algorithm, Big-O, O(n), O(n²), vectorisation, loop performance, recursion, pseudocode, divide-and-conquer"),

    (5, 3, "M3: Data & Storage", "Data Structures from a CS Perspective",
     "Arrays vs linked lists (memory layout); hash tables (how Python dicts work internally); trees (concept — how file systems and databases work); why dict lookup is O(1) but list search is O(n); stacks and queues; when each structure is appropriate; implications for pandas vs pure Python performance",
     "You use Python dicts and lists daily without knowing why dicts are O(1) for lookup while lists are O(n). This matters: if you have 1 million rows and need to find a value, storing in a dict (hash map) vs list changes your runtime from milliseconds to minutes. For spatial data: understanding tree structures explains how GIS software rapidly finds which polygon a point falls in (R-tree spatial index). You do not need to implement these — you need to know which structure to reach for.",
     "Crash Course Computer Science Ep.14 (Data Structures) | https://youtu.be/DuDz6B4cqVc | Python Data Structures (official docs) | https://docs.python.org/3/tutorial/datastructures.html",
     "Problem Solving with Algorithms and Data Structures using Python (free) | https://runestone.academy/runestone/books/published/pythonds/index.html (Ch.1-3) | Visualgo (algorithm animation) | https://visualgo.net/",
     "Benchmark exercise: compare dict vs list lookup time in Python. Compare data frame join vs nested loop join.",
     "(1) Create a Python list with 1M items and a dict with the same 1M keys; (2) Time looking up 1000 random items in each; (3) Create two pandas DataFrames with 100K rows each and merge them using pd.merge() — time it; (4) Write the same merge using nested for loops — time it; (5) Explain in 150 words why the pandas merge is so much faster (hint: it uses hash join). (6) In R: compare which(), match(), and %in% for element lookup — which is fastest and why?",
     "None",
     "array, hash table, O(1) lookup, tree, stack, queue, dict vs list, pandas internals, spatial indexing"),

    (6, 3, "M3: Data & Storage", "Version Control with Git",
     "What version control is and why researchers need it; git init, add, commit, push, pull, clone; .gitignore; branches (concept + checkout); merge conflicts (how to resolve); GitHub for researchers; commit messages as a research log; git for collaboration; undoing mistakes (revert, reset)",
     "Git is the single most important tool for research reproducibility that no methods course teaches. Every change to your analysis code is logged with a message — your commit history is a research diary. You can go back to any previous version of your analysis. You can work on two versions simultaneously using branches (e.g., one for submitting to your advisor, one for further development). For your thesis: initialise a git repository on Day 1 and commit every session. When you submit, your entire analysis history is documented.",
     "Software Carpentry: Version Control with Git (free) | https://swcarpentry.github.io/git-novice/ — work through all episodes",
     "Git for Scientists | https://neuroplausible.com/github | Happy Git with R (for R users) | https://happygitwithr.com/ | GitHub Desktop (GUI for beginners) | https://desktop.github.com/ | Oh My Git (interactive game) | https://ohmygit.org/",
     "Your own thesis project folder — initialise a git repo and practice committing your previous assignments.",
     "Set up a complete research git workflow: (1) Create a new folder, initialise a git repo; (2) Create your thesis directory structure (from Day 2 assignment); (3) Add a README.md with project description; (4) Commit with message 'Initial project structure'; (5) Create a branch called 'analysis-v1'; (6) On that branch, write a short R or Python script that loads any dataset; (7) Commit it; (8) Switch back to main and notice the script is not there; (9) Merge the branch; (10) Create a .gitignore that ignores .csv files and output/ folder.",
     "Read: Blischak, J.D., Davenport, E.R. & Wilson, G. (2016). 'A Quick Introduction to Version Control with Git and GitHub.' PLOS Computational Biology, 12(1). | https://doi.org/10.1371/journal.pcbi.1004668",
     "git init/add/commit/push/pull, .gitignore, branch, merge, revert, GitHub, commit message"),

    (7, 4, "M4: Databases & Formats", "Databases, SQL, and Data Storage Formats",
     "Relational databases: tables, primary keys, foreign keys, normalisation (1NF-3NF concept); SQL: SELECT, FROM, WHERE, GROUP BY, ORDER BY, JOIN (INNER, LEFT, RIGHT); SQLite in Python (sqlite3) and R (RSQLite); when to use databases vs CSV; data formats: CSV, JSON, Parquet, HDF5, GeoJSON; file size and performance trade-offs",
     "Census data, NFHS microdata, and administrative records are often stored in relational databases. SQL is the universal language for querying them. Even if your data arrives as CSV, knowing SQL means you can query the Census API, work with government open data portals, and collaborate with economists who use SQL databases. SQLite requires no server — it is a single file database that works in Python and R with no setup.",
     "SQLZoo (interactive SQL tutorial, free) | https://sqlzoo.net/ — complete Tutorial 1-4 | W3Schools SQL | https://www.w3schools.com/sql/",
     "Mode SQL Tutorial | https://mode.com/sql-tutorial/ | SQLite tutorial | https://www.sqlitetutorial.net/ | DBI package in R | https://dbi.r-dbi.org/ | sqlite3 in Python | https://docs.python.org/3/library/sqlite3.html",
     "Dataset: World Bank development data stored in a SQLite database. Create it yourself: download WDI CSV, load it into SQLite using Python, then query it with SQL.",
     "(1) Download WDI data as CSV; (2) Load it into SQLite using Python's sqlite3 (CREATE TABLE, INSERT); (3) Write SQL queries: (a) SELECT all rows for India; (b) SELECT top 10 countries by GDP per capita in 2020; (c) JOIN two tables (indicator data + country metadata) on country code; (d) GROUP BY region and compute average literacy rate; (4) Do the same queries in R using RSQLite and DBI; (5) Compare the SQL result to a pandas/dplyr equivalent — write 150 words on when you would choose SQL over pandas.",
     "Read: Codd, E.F. (1970). 'A Relational Model of Data for Large Shared Data Banks.' CACM, 13(6), 377-387. — the foundational paper for relational databases. Only read the abstract and first 3 pages for historical context.",
     "relational database, primary/foreign key, normalisation, SELECT/JOIN/GROUP BY, SQLite, CSV vs Parquet, DBI"),

    (8, 4, "M4: Databases & Formats", "The Research Software Ecosystem: Putting It All Together",
     "How Python, R, QGIS, STATA, and SQL fit together as a research stack; when to use which tool; APIs and web services (what an API call is at the network level); cloud computing concepts (why Colab works; what a server is); research data management (naming conventions, folder structures, metadata, README standards); open science and FAIR data principles",
     "A senior researcher's view of the toolkit: you will rarely use just one tool for a project. Python is best for data collection (APIs, scraping), large-scale processing, and spatial raster work. R is best for statistical analysis, econometrics, and publication-quality figures. QGIS is best for exploratory spatial work and cartography. STATA is the legacy standard in development economics (many replication archives are STATA-only). Knowing which tool to reach for, and how to pass data between them, is a practical superpower.",
     "FAIR Data Principles | https://www.go-fair.org/fair-principles/ | Research Data Management guide (University of Edinburgh) | https://www.ed.ac.uk/information-services/research-support/research-data-service/research-data-management",
     "The Turing Way: Research Data Management | https://the-turing-way.netlify.app/reproducible-research/rdm | Data Carpentry | https://datacarpentry.org/ | ICPSR Data Management Best Practices | https://www.icpsr.umich.edu/web/pages/datamanagement/",
     "No new dataset. Apply to your thesis project folder.",
     "Final CS capstone — set up your complete thesis research infrastructure: (1) A git-tracked project folder with proper directory structure; (2) A README.md documenting: project title, research question, data sources, software requirements, how to reproduce the analysis; (3) A data management plan (1 page): what data will you collect, how will you store it, what is your backup strategy, what can be shared publicly?; (4) A requirements.txt (Python) and sessionInfo() output (R) documenting your software versions; (5) A .gitignore that keeps raw data and credentials out of git.",
     "Read: Wilkinson, M.D. et al. (2016). 'The FAIR Guiding Principles for Scientific Data Management and Stewardship.' Scientific Data, 3, 160018. | https://doi.org/10.1038/sdata.2016.18 — now a standard citation in data management sections.",
     "research stack, tool selection, API, cloud computing, FAIR principles, data management plan, README"),
]

INTERP_DATA = [
    # ── MODULE 1: WHAT NUMBERS MEAN ───────────────────────────────────────────
    (1, 1, "M1: What Numbers Mean", "From Statistics to Substance: The Interpretation Gap",
     "Statistical significance vs substantive significance; economic/policy significance; the 'so what' test; effect size benchmarks (what is a large/small Cohen's d in practice); percentage vs percentage points; per capita vs total; units of measurement as an argument; log transformations and their interpretation",
     "The most common failure in quantitative social science is producing a correct number and then not knowing what it means. A regression coefficient of 0.03 is meaningless without knowing: (1) what the units are, (2) how that compares to the mean, (3) whether it is large enough to matter for policy. This is the 'so what' test: if your finding is true, what would change? Who would make a different decision? The failure to connect numbers to substance is why much empirical work in development economics is technically correct but policy-irrelevant.",
     "Ziliak & McCloskey 'The Cult of Statistical Significance' Ch.1-3 (library) | Gelman & Hill 'Data Analysis Using Regression' Ch.2 (interpretation sections)",
     "Ziliak & McCloskey (2008) The Cult of Statistical Significance | McCloskey's 'The Rhetoric of Economics' — on what economists actually argue | Andrew Gelman's blog (applied statistics) | https://statmodeling.stat.columbia.edu/",
     "Take any regression output from a previous assignment (any module). The dataset is your existing work.",
     "Take your IHDS regression from Econometrics Day 1 (or Statistics Day 4). For each coefficient: (1) Write the coefficient value and units; (2) Express it as a % of the mean outcome; (3) State whether the effect is large, medium, or small by field standards; (4) Write one sentence as if speaking to a policy maker: 'If we increased X by [amount], we would expect Y to change by [amount], which means...'; (5) Identify which of your coefficients is statistically significant but substantively tiny; (6) Identify which (if any) is substantively large but imprecise (wide CI). Write 300-word interpretive paragraph.",
     "Read: McCloskey, D. & Ziliak, S. (1996). 'The Standard Error of Regressions.' Journal of Economic Literature, 34(1), 97-114. | https://www.jstor.org/stable/2729411 — the paper that changed how econometricians think about significance.",
     "substantive significance, effect size benchmarks, so what test, units, percentage points vs percent, log interpretation"),

    (2, 1, "M1: What Numbers Mean", "Reading Descriptive Statistics Like a Researcher",
     "Mean vs median in skewed distributions and what each tells a different story; when to use which measure of spread; what a bimodal distribution signals; outliers as data quality problems vs substantive findings; the narrative of a histogram; comparing distributions across groups; contextualising numbers against benchmarks",
     "Descriptive statistics are not neutral summaries — they are arguments. Reporting mean income in India tells a different story than median income because the distribution is highly right-skewed. The choice of which statistic to lead with is a rhetorical choice. A bimodal distribution often signals two distinct populations (e.g., formal vs informal workers) that should not be analysed together. Learning to read descriptive statistics as a narrative — what story does this distribution tell about the underlying social process? — is what separates research from data reporting.",
     "Tukey, J.W. 'Exploratory Data Analysis' (1977) — Ch.1-3 (the classic text on reading data) | Wheelan 'Naked Statistics' Ch.2-4",
     "Data Analysis: A Bayesian Tutorial (Sivia & Skilling) for distribution thinking | Spiegelhalter 'The Art of Statistics' Ch.1-3 | Anscombe (1973) paper (already assigned in Stats module — re-read with fresh eyes here)",
     "Dataset: India Human Development Survey (IHDS) household income variable. Download if not already done: https://ihds.umd.edu/ihds-data",
     "Using IHDS income data: (1) Compute mean, median, mode, SD, P10, P25, P75, P90; (2) Plot histogram and box plot; (3) Write separate one-sentence interpretations of mean and median that tell different stories about Indian household incomes; (4) Compute the ratio of P90 to P10 — what does this tell you that SD does not?; (5) Separate by rural/urban and compare distributions — write 200 words describing what a social scientist would conclude from comparing these two distributions; (6) Identify any outliers — are they data errors or real observations?",
     "Read (reread): Anscombe, F.J. (1973). 'Graphs in Statistical Analysis.' American Statistician, 27(1), 17-21. | This time, focus on: what does Anscombe demonstrate about the inadequacy of summary statistics alone?",
     "mean vs median narrative, bimodal distributions, outliers as signal, P90/P10 ratio, distributional comparison, histogram reading"),

    (3, 2, "M2: Interpreting Models", "Interpreting Regression Coefficients: A Full Taxonomy",
     "Continuous predictor (linear); log-linear, linear-log, log-log models and their interpretations; binary predictor (dummy variable); categorical predictor (reference category and its implications); interaction terms (what the coefficient does and does not mean); fixed effects coefficients (what they absorb and what they do not); model intercept interpretation; marginal effects vs elasticities",
     "Most students can run a regression. Very few can interpret every coefficient correctly. The interpretation depends entirely on the functional form: in a log-linear model, the coefficient is approximately a percentage change. In a log-log model, the coefficient is an elasticity. An interaction term coefficient does not tell you the total effect of either variable — it tells you the difference in slopes. Fixed effects coefficients absorb all between-unit variation — the remaining coefficient is a within-unit effect. Getting these wrong produces incorrect policy conclusions even from correct code.",
     "Wooldridge 'Introductory Econometrics' Ch.2 (pp. 35-50 on log models) + Ch.6 (multiple regression interpretation) | Kennedy 'A Guide to Econometrics' Ch.3 (the most practical interpretation guide)",
     "UCLA IDRE Regression Interpretation | https://stats.oarc.ucla.edu/other/mult-pkg/introduction-to-linear-mixed-models/ | Interpretation of log models | https://stats.oarc.ucla.edu/other/mult-pkg/faq/general/faqhow-do-i-interpret-a-regression-model-when-some-variables-are-log-transformed/ | Kennedy 'A Guide to Econometrics' (library)",
     "Dataset: IHDS regression from your previous work. You are re-interpreting, not re-running.",
     "For your IHDS regression output: (1) Write a formal interpretation of every coefficient in one sentence each, specifying units; (2) Re-run with log(income) as outcome — re-interpret all coefficients; (3) Add an interaction between education and SC/ST dummy — write out the full interpretation: 'For SC/ST individuals, a one-unit increase in education is associated with a change of [β_edu + β_interaction] in log income, compared to [β_edu] for non-SC/ST individuals'; (4) Create a table with three columns: Variable | Coefficient | Plain-English Interpretation. This table is the template for your thesis results section.",
     "Read: King, G., Tomz, M. & Wittenberg, J. (2000). 'Making the Most of Statistical Analyses.' American Journal of Political Science, 44(2), 347-361. | https://www.jstor.org/stable/2669316 — on interpreting and presenting statistical results for maximum substantive clarity.",
     "log-linear interpretation, log-log elasticity, dummy variable, interaction term interpretation, reference category, fixed effects absorption, marginal effects"),

    (4, 2, "M2: Interpreting Models", "Interpreting Causal Claims: What Your Estimate Identifies",
     "The counterfactual question; what OLS identifies (correlation with controls); what IV identifies (LATE — the Local Average Treatment Effect); what DiD identifies (ATT — the Average Treatment Effect on the Treated); what RDD identifies (LATE at the cutoff); internal vs external validity; SUTVA (Stable Unit Treatment Value Assumption); spillovers and their consequences; generalising from a local estimate",
     "The most important question to ask about any causal estimate is: for whom is this effect identified, and under what conditions? IV gives you the LATE — the effect for compliers (those who change treatment status because of the instrument). This may be a small and unusual subgroup. RDD gives you the effect at the cutoff — which may not be representative of the broader population. DiD gives you the ATT — the effect on the treated group, not on everyone. Confusing these leads to policy recommendations that generalise beyond what the data supports.",
     "Angrist & Pischke 'Mostly Harmless Econometrics' Ch.1-2 (what is causal inference?) | Imbens & Rubin 'Causal Inference for Statistics' Ch.1 (free preview) | https://www.cambridge.org/core/books/causal-inference-for-statistics-social-and-biomedical-sciences/71126BE90C58F1A431FE9B185943F2DA",
     "The Effect Book: Causal Inference Introduction | https://theeffectbook.net/ch-CausalPaths.html | Pearl & Mackenzie 'The Book of Why' Ch.1-3 (for intuition, not formalism) | Scott Cunningham's Mixtape: Potential Outcomes | https://mixtape.scunning.com/04-potential_outcomes",
     "Dataset: Re-use any previous causal exercise (IV from Econometrics Day 2, or DiD from Day 4).",
     "For your chosen causal paper or exercise: (1) State the counterfactual question precisely: 'Compared to what?'; (2) Identify what the estimator identifies (LATE, ATT, etc.) and write who exactly is in that group; (3) State one threat to internal validity with a specific example from your context; (4) State one threat to external validity: to what population does this estimate generalise?; (5) Write a 200-word 'Causal Interpretation' paragraph as it would appear in a paper, using appropriately hedged language; (6) Write the same finding for a policy brief audience in 100 words — notice how different the language needs to be.",
     "Read: Deaton, A. & Cartwright, N. (2018). 'Understanding and Misunderstanding Randomized Controlled Trials.' Social Science & Medicine, 210, 2-21. | https://doi.org/10.1016/j.socscimed.2017.12.005 — a rigorous critique of the external validity of experimental estimates.",
     "counterfactual, LATE, ATT, SUTVA, internal validity, external validity, complier, generalisation"),

    (5, 3, "M3: Interpreting Space", "Interpreting Spatial Patterns: What a Map Is and Is Not Saying",
     "Maps as arguments, not neutral representations; the Modifiable Areal Unit Problem (MAUP); ecological fallacy; spatial clustering vs causal explanation; what a Moran's I value means substantively; LISA map interpretation; spatial spillovers (what they mean in practice); scale sensitivity; reading maps critically",
     "A map that shows high e-waste activity concentrated in low-income areas is not evidence that e-waste causes poverty, or that poverty causes e-waste concentration. It is a description of a spatial pattern that demands a theoretical explanation. The MAUP is a fundamental problem: the same underlying data produces different patterns depending on how you draw the boundaries (wards vs districts vs zones). The ecological fallacy is using aggregate-level correlations to make individual-level claims. These are not optional caveats — they are the conditions under which any spatial finding is scientifically defensible.",
     "Openshaw, S. (1983). 'The Modifiable Areal Unit Problem.' CATMOG 38 | Anselin, L. 'Spatial Econometrics' Ch.1 (conceptual sections) | Robinson, W.S. (1950) 'Ecological Correlations and the Behavior of Individuals' — the original ecological fallacy paper",
     "Fotheringham & Wong (1991) 'The Modifiable Areal Unit Problem in Multivariate Statistical Analysis' | GIS Geography: MAUP | https://gisgeography.com/maup-modifiable-areal-unit-problem/ | Openshaw (1983) CATMOG | https://www.qmrg.org.uk/catmog/",
     "Dataset: SHRUG district + sub-district data (if available). Compare your LISA map at different spatial scales.",
     "Demonstrate MAUP empirically: (1) Take the nighttime lights data you have and aggregate to (a) sub-district, (b) district, (c) state level; (2) Run Moran's I at each level — does the spatial autocorrelation statistic change?; (3) Make LISA maps at each scale — do the clusters shift?; (4) Write 200 words: 'At which spatial scale is the analysis most appropriate for my research question, and why?'; (5) Find one published spatial paper and identify: did the authors address the MAUP? Did they address the ecological fallacy? Write 100-word critical note.",
     "Read: Robinson, W.S. (1950). 'Ecological Correlations and the Behavior of Individuals.' American Sociological Review, 15(3), 351-357. — the original ecological fallacy paper. Short and landmark.",
     "MAUP, ecological fallacy, spatial clustering vs causation, scale sensitivity, LISA interpretation, spatial spillover meaning"),

    (6, 3, "M3: Interpreting Space", "Communicating Uncertainty: Confidence Intervals, Error Bars, and What They Mean",
     "What a confidence interval actually means (and the three most common wrong interpretations); credible intervals (Bayesian) vs CI (frequentist); how to describe uncertainty in plain English; forest plots and coefficient plots (ggplot2/dotwhisker); the difference between 'not significant' and 'no effect'; communicating model uncertainty to non-technical audiences; why wide CIs are informative",
     "The most common misinterpretation in all of empirical social science: 'there is a 95% probability the true value lies in this interval.' This is wrong — in frequentist statistics, the true value is fixed; the interval is random. What a 95% CI means: if you repeated this sampling procedure 100 times, approximately 95 of the resulting intervals would contain the true value. A non-significant result does not mean the effect is zero — it means the data are insufficient to distinguish the effect from zero. These distinctions matter enormously when writing up research.",
     "Cumming, G. 'The New Statistics' | https://www.ncbi.nlm.nih.gov/pmc/articles/PMC3955264/ | Morey et al. (2016) 'The Fallacy of Placing Confidence in Confidence Intervals' | https://link.springer.com/article/10.3758/s13423-015-0947-8",
     "dotwhisker package for R coefficient plots | https://cran.r-project.org/web/packages/dotwhisker/ | ggplot2 for error bars | Understanding CIs — StatQuest | https://youtu.be/TqOeMYtOc1w | Rethinking 'Statistical Significance' — Nature (2019) | https://www.nature.com/articles/d41586-019-00857-9",
     "Dataset: Your regression output from any previous assignment. You are visualising and re-describing, not re-running.",
     "(1) From your IHDS regression, plot a coefficient plot using dotwhisker or ggplot2 with 95% CI error bars; (2) Identify coefficients where the CI is very wide — write a sentence explaining what this means substantively ('We cannot rule out effects ranging from X to Y'); (3) Identify a coefficient that is not statistically significant but has a large point estimate — write 100 words explaining why reporting it as 'no effect' would be misleading; (4) Write two versions of a CI interpretation: one technically correct, one common-but-wrong; (5) Present your coefficient plot to a non-technical friend and explain each CI in one sentence.",
     "Read: Greenland, S. et al. (2016). 'Statistical Tests, P Values, Confidence Intervals, and Power: A Guide to Misinterpretations.' European Journal of Epidemiology, 31(4), 337-350. | https://doi.org/10.1007/s10654-016-0149-3",
     "CI interpretation, credible interval, not significant ≠ no effect, coefficient plot, forest plot, uncertainty communication"),

    (7, 4, "M4: Writing it Up", "Turning Results into Research Writing: The Results Section",
     "Structure of a results section; the claim-evidence-interpretation (CEI) structure for each finding; hedging language in academic writing ('is associated with', 'we find evidence that', 'consistent with'); what goes in a results section vs discussion; reporting conventions (APA, economics style); integrating tables and figures into the text; writing for a non-specialist reader",
     "The results section is not a list of numbers — it is a structured argument. Each finding is: (1) a claim, (2) supported by evidence (the statistic), (3) interpreted for its meaning. 'Table 2 shows a coefficient of 0.23' is not a results section — it is a table description. A results section says: 'We find a positive association between years of education and household income (β = 0.23, SE = 0.04, p < 0.001), suggesting that each additional year of schooling is associated with approximately 23% higher income, holding other household characteristics constant.' Note: the hedge ('suggesting'), the mechanism ('each additional year'), and the condition ('holding constant').",
     "Sword, H. 'Stylish Academic Writing' (library) | Silvia, P. 'How to Write a Lot' (for productivity) | Becker, H. 'Writing for Social Scientists'",
     "The Elements of Style (Strunk & White) — brevity and clarity | American Economic Review Style Guide | https://www.aeaweb.org/journals/aer/submissions/accepted-articles/styleguide | Writing a Results Section (Purdue OWL) | https://owl.purdue.edu/owl/research_and_citation/apa_style/apa_formatting_and_style_guide/apa_sample_paper.html",
     "Your own regression output from any previous assignment.",
     "Take your IHDS or NFHS regression (whichever is most developed). Write a complete results section of 400-500 words that: (1) opens with the most important finding; (2) presents each coefficient using CEI structure; (3) integrates at least one table and one figure as evidence; (4) uses appropriately hedged language throughout; (5) ends with a sentence connecting the findings to your research question. Then exchange with a classmate (or read it aloud to yourself) — does every sentence make a clear claim?",
     "Read: Belcher, W.L. (2009). 'Writing Your Journal Article in Twelve Weeks.' Ch.7 (Results section) — extremely practical guidance on academic writing from a social scientist's perspective.",
     "CEI structure, hedging language, claim vs evidence, results vs discussion, regression reporting conventions, integrating tables and figures"),

    (8, 4, "M4: Writing it Up", "Critical Reading of Empirical Papers",
     "How to read a methods section: identifying the identification strategy; spotting threats to validity; evaluating data quality; understanding robustness checks; the anatomy of a referee report; common methodological weaknesses in development economics; replication as a reading strategy; the difference between peer-reviewed and working paper quality",
     "Reading a methods section critically is the mirror of writing one. The questions to ask: What is the unit of analysis? What is the treatment and control? What is the identification assumption and is it stated? What data quality issues exist? What is the attrition rate? Did they test pre-trends? Did they address MAUP? Are their standard errors clustered appropriately? Building this checklist takes years in a seminar room — this day is a shortcut. The ability to read papers critically protects you from citing flawed evidence in your own work.",
     "Angrist & Pischke 'Mostly Harmless Econometrics' Preface (what makes an empirical paper credible) | Reviewer guidelines for top journals (NBER, AER) | Critical reading framework (Morgan & Winship 'Counterfactuals and Causal Inference' Ch.1)",
     "How to read an econometrics paper — various guides | Journal of Development Economics author guidelines | https://www.journals.elsevier.com/journal-of-development-economics | NBER Working Paper series | https://www.nber.org/papers",
     "Choose any published paper with India data from: Journal of Development Economics, World Development, Economic and Political Weekly (EPW), or AER. Suggested: any paper from J-PAL's evidence repository | https://www.povertyactionlab.org/evidence-to-policy",
     "Select one paper from J-PAL or AEA with India data. Answer in writing: (1) What is the research question? (2) What is the identification strategy? (3) What is the key identifying assumption — is it stated? Is it plausible? (4) What data are used — quality, coverage, potential biases? (5) Are the standard errors appropriate (clustered? heteroskedasticity-robust?)? (6) What robustness checks do they run — are they convincing? (7) What is the main threat to internal validity they don't fully address? (8) Write a 200-word mock referee comment raising your most serious concern.",
     "Read: Hamermesh, D.S. (2007). 'Viewpoint: Replication in Economics.' Canadian Journal of Economics, 40(3), 715-733. | https://doi.org/10.1111/j.1365-2966.2007.00428.x — on why replication matters and what it reveals.",
     "identification strategy, identifying assumption, data quality, robustness checks, referee report, threats to validity, pre-trends"),

    (9, 5, "M5: Visual Interpretation", "Data Visualisation as Interpretation",
     "When visualisations mislead: truncated y-axes; cherry-picked time periods; dual axes creating spurious correlation; inappropriate smoothing; comparing non-comparable groups without normalisation; the distinction between visualising data and visualising model results; effective use of uncertainty bands; choosing chart type to match the claim being made",
     "Every chart makes a claim. A bar chart with a y-axis starting at 90 makes a 2% difference look like a 50% difference. Plotting two variables with dual axes can make uncorrelated series look correlated. These are not just aesthetic issues — they are epistemological ones. The chart literally changes what the viewer believes. Learning to identify these manipulations in others' work, and avoid them in your own, is a core research integrity skill. The test: does the visual representation match the statistical reality?",
     "Wilke, C. 'Fundamentals of Data Visualization' (free) | https://clauswilke.com/dataviz/ | Cairo, A. 'How Charts Lie' (highly readable) | Tufte, E. 'The Visual Display of Quantitative Information'",
     "Calling Bullshit (University of Washington course) | https://www.callingbullshit.org/ | Flowing Data | https://flowingdata.com/ | WTF Visualizations | https://viz.wtf/ | Junk Charts | https://junkcharts.typepad.com/",
     "Dataset: Find 3 misleading charts in the wild (news websites, government reports, social media). Print them or screenshot them.",
     "(1) Find 3 misleading charts online (government statistics, news, NGO reports — India context preferred); (2) For each: identify the specific misleading technique; (3) Download the underlying data and re-make the chart correctly; (4) Write a 100-word 'correction note' for each; (5) Take one of your own charts from a previous assignment — apply Tufte's data-ink ratio principle and simplify it; (6) Compare before/after and write 100 words on what the cleaner version communicates more effectively.",
     "Read: Huff, D. (1954). 'How to Lie with Statistics.' — still the best short book on statistical manipulation. Read the whole thing (it is 144 pages, very fast). Available in many libraries.",
     "truncated axes, dual axes, cherry-picking, inappropriate smoothing, data-ink ratio, visualisation as argument, uncertainty bands"),

    (10, 5, "M5: Visual Interpretation", "From Data to Policy Argument: The Full Interpretive Chain",
     "The research-to-policy pipeline: data → analysis → finding → interpretation → policy implication; writing for different audiences (academic paper vs policy brief vs public); quantitative findings embedded in qualitative context; what makes a finding policy-relevant; the difference between 'statistically significant' and 'actionable'; practising the translation exercise across audiences",
     "The full chain: you have run your spatial regression, produced your LISA map, and found that informal e-waste activity is spatially clustered in low-income wards with poor formal waste infrastructure. Now what? A policy implication is not just 'government should do something' — it is a specific, falsifiable, actionable recommendation that follows from your specific finding. The connection between the statistical finding and the policy recommendation requires a chain of reasoning that the researcher must make explicit. This chain is what separates research from advocacy.",
     "Cartwright, N. & Hardie, J. 'Evidence-Based Policy: A Practical Guide to Doing It Better' (Oxford, library) | Banerjee & Duflo 'Poor Economics' Ch.1 (model of research-to-policy) | ODI Research-to-Policy toolkit | https://odi.org/en/publications/",
     "ODI: How to write a policy brief | https://odi.org/en/publications/how-to-write-a-policy-brief/ | IZA policy brief format | https://www.iza.org/publications/pp | CPPR policy brief examples | J-PAL policy publications | https://www.povertyactionlab.org/policy-publications",
     "Dataset: Your most developed previous assignment — the IHDS/NFHS regression or the spatial analysis.",
     "Full translation exercise. Take your best previous result. Write three versions: (1) Academic: a 400-word results + discussion section in journal style with citations, hedge language, caveats; (2) Policy brief: a 300-word brief for a government ministry official with no statistics training — key finding, what it means, what you recommend, what you are uncertain about; (3) Public communication: a 200-word summary for a newspaper op-ed. Read all three aloud. Notice what changes and what must stay the same. Write 150 words reflecting on what was lost and gained in each translation.",
     "Read: Cairney, P. (2016). 'The Politics of Evidence-Based Policy Making.' Palgrave. Ch.1 only (30 pages). — on the gap between what researchers find and what policy makers do with it.",
     "research-to-policy chain, policy brief format, audience translation, actionable recommendation, quantitative-qualitative integration, policy relevance test"),
]

PYTHON_DATA = [
    # ── MODULE 1: FOUNDATIONS ─────────────────────────────────────────────────
    (1, 1, "M1: Foundations", "Setup + Python Basics",
     "Anaconda install; Jupyter notebooks; variables; data types (int, float, str, bool); arithmetic operators; f-strings; print()",
     "Open Anaconda Navigator → launch Jupyter. Write your first notebook: store your name and age as variables, compute age in months, print a formatted sentence. Understand: Python is interpreted line by line. Int vs float matters for division. Strings are sequences.",
     "Python.org Official Tutorial §§1–3 | https://docs.python.org/3/tutorial/",
     "Automate the Boring Stuff Ch.1 (free) | https://automatetheboringstuff.com/2e/chapter1/ | YouTube: Corey Schafer 'Python Beginner Tutorial' Ep1-2 | https://youtu.be/YYXdXT2l-Gg",
     "No external dataset yet. Practice: https://www.w3schools.com/python/exercise.asp (Exercises 1-10)",
     "Write a Jupyter notebook that: (1) stores India's GDP (2023), population, area as variables; (2) computes GDP per capita and population density; (3) prints a 3-sentence summary using f-strings.",
     "None yet — focus on getting environment working.",
     "Variables, types, operators, f-strings, Jupyter cells"),

    (2, 1, "M1: Foundations", "Control Flow + Functions",
     "if/elif/else; for loops; while loops; list comprehensions; def keyword; arguments, return values; scope; lambda functions",
     "Control flow is the brain of any program. Loops let you repeat operations across datasets — this is how you process 1000 survey rows without copy-pasting. Functions package reusable logic. Understand: indentation IS syntax in Python, not cosmetic.",
     "Python.org Tutorial §§4–5 | https://docs.python.org/3/tutorial/controlflow.html",
     "Automate the Boring Stuff Ch.2-3 | https://automatetheboringstuff.com/ | Corey Schafer Ep3-5 | https://youtu.be/DZwmZ8Usvnk",
     "Practice: HackerRank Python (Basic) - first 10 challenges | https://www.hackerrank.com/domains/python",
     "Write a function that takes a list of numbers (representing district-level poverty rates) and returns: mean, median, min, max, and number of districts above the national average. Test with data: [28.3, 14.2, 45.1, 33.6, 9.8, 52.4, 22.1, 18.7]",
     "None",
     "if/elif/else, for loops, list comprehensions, def, return, scope"),

    (3, 1, "M1: Foundations", "Data Structures",
     "Lists (ordered, mutable); Tuples (immutable); Dictionaries (key-value); Sets (unique); nested structures; slicing; common methods (.append, .pop, .keys, .values, .items)",
     "Dictionaries are crucial for research data — think of each respondent as a dict: {name: 'Asha', age: 24, district: 'Chennai', income: 12000}. Lists hold sequences. Knowing which structure to use is a core skill. Practice slicing — it comes up constantly in data analysis.",
     "Real Python - Python Data Structures | https://realpython.com/python-data-structures/",
     "Python.org Tutorial §§5 | https://docs.python.org/3/tutorial/datastructures.html | Corey Schafer Ep6-7",
     "Practice problem: https://www.practicepython.org/ - Exercises 1, 4, 5, 9, 14",
     "Represent 5 households from a survey as a list of dictionaries. Each dict has: household_id, district, income, num_members, has_toilet (bool). Write a loop that (a) prints households with income < 15000; (b) computes average income; (c) counts how many lack toilet access.",
     "None",
     "list, dict, tuple, set, slicing, nested structures"),

    (4, 1, "M1: Foundations", "NumPy — Numerical Computing",
     "ndarray vs list; array creation (zeros, ones, arange, linspace); indexing and slicing; broadcasting; vectorised operations; shape/reshape; statistical functions (mean, std, median, percentile); random module",
     "NumPy is the engine under everything else — pandas, scipy, scikit-learn all sit on top of it. Key insight: NumPy operations run on entire arrays at once (vectorised), not element by element like a loop. This makes them 100x faster on large datasets. You will use np.mean(), np.std(), np.percentile() constantly in research.",
     "NumPy Official Quickstart | https://numpy.org/doc/stable/user/quickstart.html",
     "NumPy for Absolute Beginners (official) | https://numpy.org/doc/stable/user/absolute_beginners.html | Corey Schafer NumPy tutorial | https://youtu.be/GB9ByFAIAH4",
     "Dataset: Use NumPy to work with NFHS-5 state-level summary data (download from https://rchiips.org/nfhs/NFHS-5Reports/NationalFactSheet_India.pdf — extract numbers manually, ~5 states). Practice loading arrays, computing descriptive stats.",
     "Using NumPy: (1) Create an array of 50 random 'household income' values between 5000-80000; (2) compute mean, median, std, 25th and 75th percentile; (3) create a boolean mask for households below median; (4) compute what fraction is below median. No pandas — raw NumPy only.",
     "None",
     "ndarray, broadcasting, vectorisation, indexing, shape, np.mean/std/percentile"),

    (5, 2, "M2: Pandas — Data Analysis", "Pandas Part 1: Loading and Exploring Data",
     "Series vs DataFrame; pd.read_csv(), pd.read_excel(); .head(), .tail(), .info(), .describe(), .shape, .dtypes; index and column selection; .iloc[] vs .loc[]; value_counts(); .unique()",
     "Pandas is your primary tool for 90% of quantitative research work. A DataFrame is an Excel sheet inside Python — rows are observations, columns are variables. Master the distinction between .loc (label-based) and .iloc (position-based) — getting this wrong causes silent errors. .describe() gives you a first diagnostic of any dataset.",
     "Python for Data Analysis (Wes McKinney, 3rd ed.) Ch.5 | Free via your library | https://wesmckinney.com/book/",
     "Pandas official getting started | https://pandas.pydata.org/docs/getting_started/intro_tutorials/ | Corey Schafer Pandas series | https://youtu.be/ZyhVh-qRZPA",
     "Dataset: Titanic (classic ML dataset, perfect for learning pandas) | https://www.kaggle.com/datasets/yasserh/titanic-dataset — download titanic.csv",
     "Load titanic.csv. Answer: (1) How many passengers? (2) What fraction survived? (3) What is the mean age? (4) How many are in each passenger class? (5) How many missing values in 'Age'? Write each answer as a comment next to the code that produces it.",
     "None",
     "DataFrame, Series, read_csv, loc, iloc, describe, value_counts, dtypes"),

    (6, 2, "M2: Pandas — Data Analysis", "Pandas Part 2: Cleaning and Transformation",
     "Handling missing values (.isna(), .fillna(), .dropna()); renaming columns; changing data types; str accessor; .apply(); lambda with apply; merging/joining (pd.merge, left/right/inner/outer); .groupby() + .agg(); .pivot_table()",
     "Real-world data is always dirty. India Census data has spelling inconsistencies, NFHS data has coded values that need decoding, admin data has different state names across datasets. Merging is the single most powerful skill — it lets you combine household survey data with district-level administrative data, or spatial boundaries with census attributes. groupby is your workhorse for sub-group analysis.",
     "Python for Data Analysis Ch.7-8 | https://wesmckinney.com/book/data-cleaning",
     "Pandas Merging Tutorial | https://pandas.pydata.org/docs/user_guide/merging.html | Towards Data Science: '8 ways to merge pandas DataFrames' | Keith Galli pandas tutorial | https://youtu.be/vmEHCJofslg",
     "Dataset: World Bank World Development Indicators | https://databank.worldbank.org/source/world-development-indicators — download GDP per capita, literacy, infant mortality for 50 countries, 2000-2022. This is a real panel dataset with missing values.",
     "With WDI data: (1) Identify all columns with >20% missing values; (2) fill numeric missing with column median; (3) rename columns to snake_case; (4) merge with a second table of regional classifications (download from World Bank); (5) compute mean GDP per capita by region and year using groupby.",
     "None",
     "isna, fillna, merge, groupby, agg, pivot_table, apply, str accessor"),

    (7, 2, "M2: Pandas — Data Analysis", "Data Visualisation: Matplotlib & Seaborn",
     "Matplotlib: figure/axes architecture; plt.plot(), plt.scatter(), plt.bar(), plt.hist(), plt.subplots(); customisation (labels, titles, colors, legends). Seaborn: sns.histplot(), sns.boxplot(), sns.scatterplot(), sns.heatmap(), sns.pairplot(); FacetGrid for multi-panel plots",
     "Publication-quality charts require understanding the figure/axes object hierarchy in Matplotlib. Always use the object-oriented interface (fig, ax = plt.subplots()) not the procedural plt. interface — it breaks in loops and multi-panel layouts. Seaborn is built on Matplotlib but handles statistical aesthetics automatically. Learn to make your charts tell one clear story.",
     "Matplotlib official tutorials | https://matplotlib.org/stable/tutorials/index.html | Seaborn tutorial | https://seaborn.pydata.org/tutorial.html",
     "Python Graph Gallery (code for every chart type) | https://python-graph-gallery.com/ | Corey Schafer Matplotlib | https://youtu.be/UO98lJQ3QGI | Seaborn crash course | https://youtu.be/6GUZXDef2U0",
     "Dataset: India district-level development data | https://shrug-data.org/ (SHRUG — Social and Human Sciences Research Using GIS) — download district-level data",
     "Using any India dataset: make (1) histogram of income distribution; (2) boxplot comparing a variable across 5 states; (3) scatter plot of literacy vs infant mortality with regression line; (4) heatmap of correlation matrix; (5) 4-panel figure combining all of the above. Save as 300dpi PNG for a 'paper figure'.",
     "None",
     "fig/ax architecture, subplots, seaborn aesthetics, FacetGrid, saving figures"),

    (8, 2, "M2: Pandas — Data Analysis", "Data Cleaning: A Real Workflow",
     "Duplicates; outlier detection (IQR, z-scores); string normalisation (whitespace, case, encoding); date parsing; pd.Categorical; binning with pd.cut() and pd.qcut(); encoding categoricals (get_dummies); pipelines with method chaining",
     "A senior researcher's rule of thumb: plan for 60-70% of your time to be data cleaning. The most common errors — merging on mismatched state names, treating coded values as numbers, ignoring encoding errors on Indian-language fields — happen at this stage. Build a systematic cleaning checklist and run it on every new dataset as a ritual.",
     "Towards Data Science: 'The Ultimate Guide to Data Cleaning' | https://towardsdatascience.com/the-ultimate-guide-to-data-cleaning-3969843991d4",
     "Python for Data Analysis Ch.7 | https://wesmckinney.com/book/ | Real Python: 'Pandas Data Cleaning' | https://realpython.com/python-data-cleaning-numpy-pandas/",
     "Dataset: Messy real-world dataset — DHS/NFHS district summary (download from https://rchiips.org/nfhs/) which has inconsistent district names, coded values, and mixed types.",
     "Take a raw NFHS district file: (1) identify and remove duplicates; (2) detect outliers using IQR method; (3) normalise state names (strip whitespace, fix capitalisation); (4) create 5 income quintile bins with pd.qcut(); (5) document your cleaning steps in markdown cells as if writing a data appendix for a paper.",
     "None",
     "duplicates, outliers, IQR, z-score, pd.cut, pd.qcut, get_dummies, method chaining"),

    (9, 3, "M3: Python for Research", "APIs, Web Data, and Automation",
     "HTTP requests with requests library; JSON parsing; REST API basics; World Bank API, UN Data API; simple web scraping with BeautifulSoup; reading PDFs with pdfplumber; automating repetitive tasks",
     "Research increasingly requires pulling data from live databases. The World Bank, UN, OECD, Census all have APIs. Understanding how to query them means you never have to manually download Excel files again. BeautifulSoup lets you extract tables from government websites that don't offer downloads. This is a practical superpower for Indian policy research.",
     "Requests library docs | https://requests.readthedocs.io/ | World Bank API docs | https://datahelpdesk.worldbank.org/knowledgebase/topics/125589",
     "Real Python: 'Python Requests Tutorial' | https://realpython.com/python-requests/ | Beautiful Soup docs | https://www.crummy.com/software/BeautifulSoup/bs4/doc/ | wbdata Python package | https://github.com/oliversherouse/wbdata",
     "Live API: World Bank API | https://api.worldbank.org/v2/country/IN/indicator/SP.POP.TOTL?format=json — pull India population time series. Also try: https://api.worldbank.org/v2/country/IN/indicator/NY.GDP.PCAP.CD?format=json for GDP per capita.",
     "Using the World Bank API (via wbdata or requests): download GDP per capita, poverty headcount, literacy rate, and infant mortality for all South Asian countries (1990-2023). Clean the result into a tidy panel DataFrame and save as CSV. Bonus: pull the same data for BRICS countries and compare.",
     "None",
     "HTTP requests, JSON, REST API, BeautifulSoup, automation, wbdata"),

    (10, 3, "M3: Python for Research", "Scipy for Research + Statistical Tests",
     "scipy.stats module: t-test (ttest_ind, ttest_rel); chi-square test; Mann-Whitney U; Kolmogorov-Smirnov; ANOVA (f_oneway); correlation (pearsonr, spearmanr); confidence intervals; effect sizes",
     "Statistical tests are the grammar of quantitative social science claims. Knowing which test to apply and why (t-test for continuous outcomes between two groups; chi-square for categorical associations; Mann-Whitney when normality cannot be assumed) is as important as knowing how to run it. Always report effect size alongside p-value — a statistically significant result with tiny effect size is usually uninteresting for policy.",
     "SciPy Statistical Functions | https://docs.scipy.org/doc/scipy/reference/stats.html",
     "StatQuest 'Statistics Fundamentals' playlist | https://youtube.com/playlist?list=PLblh5JKOoLUK0FLuzwntyYI10UQFUhsY9 | Real Python SciPy tutorial | https://realpython.com/python-scipy-cluster-optimize/",
     "Dataset: NFHS-5 vs NFHS-4 state comparison data | https://rchiips.org/nfhs/ — test whether improvement in key indicators between NFHS-4 and NFHS-5 is statistically significant.",
     "Using NFHS data: (1) Test if mean child stunting rate is significantly different between North and South India (t-test); (2) Test if toilet access and child mortality are associated (chi-square); (3) Test if the distribution of anaemia rates is normal (KS test); (4) Compute 95% CI for mean under-5 mortality. Report results in a table as you would in a paper appendix.",
     "None",
     "t-test, chi-square, Mann-Whitney, ANOVA, pearsonr, confidence intervals, effect size"),

    (11, 3, "M3: Python for Research", "Introduction to Scikit-learn (for Social Scientists)",
     "Train/test split; cross-validation; linear regression with sklearn; Ridge and Lasso regularisation; Random Forest for feature importance (variable selection); sklearn pipeline; model evaluation metrics (R², RMSE, MAE)",
     "Scikit-learn matters for social scientists not primarily for prediction but for two things: (1) variable selection via regularisation (Lasso sets irrelevant coefficients to zero — useful when you have 50 indicators and need to identify the most important); (2) feature importance from tree models to understand which district-level variables best predict a social outcome. Always distinguish: you are doing inference, not prediction.",
     "Scikit-learn User Guide | https://scikit-learn.org/stable/user_guide.html | Introduction to Statistical Learning (ISLR) Ch.6 | https://www.statlearning.com/ (free PDF)",
     "Towards Data Science: 'Sklearn for Social Science' | StatQuest 'Lasso Regression' | https://youtu.be/NGf0voTMlcs | ISLR videos | https://www.dataschool.io/15-hours-of-expert-machine-learning-videos/",
     "Dataset: India Human Development Survey (IHDS-II) | https://www.icpsr.umich.edu/web/DSDR/studies/36151 — predict household income from education, land, caste, state variables using Lasso.",
     "Using IHDS data: (1) Build OLS regression of log income on 8 predictors; (2) Build Lasso version — which variables get zeroed out?; (3) Build Random Forest and plot feature importance; (4) Compare R² across models in a table. Write a 200-word paragraph interpreting the findings as if writing a results section.",
     "None",
     "train_test_split, cross-validation, Lasso, Ridge, feature importance, R², RMSE"),

    (12, 4, "M4: Capstone", "Python Capstone Mini-Project",
     "End-to-end research workflow: data acquisition → cleaning → analysis → visualisation → brief write-up. Full pipeline with real social science data.",
     "This day is synthesis, not new content. Build a complete mini-project from scratch. Choose one of the three project options. The goal is experiencing the entire research data pipeline as it actually works — messiness, dead ends, and all. Document every step in Jupyter markdown as if writing a methods section.",
     "None — all previous resources apply.",
     "Review all previous resources. Kaggle kernels for workflow inspiration | https://www.kaggle.com/code",
     "Project A: SHRUG data — map district-level development across India, test for North-South divide using t-test, visualise with 4-panel chart. | Project B: WDI panel — test convergence hypothesis (do poorer countries grow faster?) using OLS. | Project C: NFHS data — build a 5-indicator deprivation index, identify top-10 and bottom-10 districts. Dataset links provided in Days 6,7,8.",
     "Complete your chosen project. Deliverable: Jupyter notebook with markdown commentary throughout, 3+ visualisations, one summary table, and a 300-word 'findings' section.",
     "Read: Gentzkow & Shapiro (2014) 'Code and Data for the Social Sciences: A Practitioner's Guide' | https://web.stanford.edu/~gentzkow/research/CodeAndData.pdf — the definitive guide to reproducible research workflows.",
     "Full pipeline, reproducibility, clean code, research communication"),

    # ── MODULE 5: Advanced Python for Research ───────────────────────────────
    (13, 4, "M5: Advanced Python", "Text Analysis with Python (for Policy Documents)",
     "NLTK and spaCy basics; tokenisation; stopword removal; frequency analysis; TF-IDF; basic sentiment; reading PDFs programmatically; analysing policy document corpora",
     "Your qualitative research involves policy documents (E-Waste Rules, NFHS reports, government circulars). Python lets you process hundreds of documents systematically — finding which terms co-occur, how frequently 'informal' vs 'formal' appears across policy eras, or tracking how 'circularity' entered environmental policy discourse. This bridges your qualitative expertise with computational methods.",
     "NLTK Book (free) | https://www.nltk.org/book/ | spaCy 101 | https://spacy.io/usage/spacy-101",
     "Towards Data Science: 'NLP for Social Scientists' | scikit-learn text feature extraction | https://scikit-learn.org/stable/modules/feature_extraction.html#text-feature-extraction",
     "Dataset: E-Waste Management Rules 2016 + 2022 Amendment (download from MoEFCC website | https://moef.gov.in) + WHO reports on e-waste. Compute word frequencies and compare across document versions.",
     "Download India's E-Waste Rules 2016 and 2022 Amendment PDFs. Extract text with pdfplumber. Compute: (1) top-30 words excluding stopwords in each version; (2) TF-IDF scores to find words uniquely important to each version; (3) frequency of terms 'informal', 'formal', 'producer', 'consumer', 'circular'. Write 200-word interpretation.",
     "Read: Grimmer, J. & Stewart, B.M. (2013). 'Text as Data.' Political Analysis, 21(3), 267-297. | https://www.cambridge.org/core/journals/political-analysis/article/text-as-data-the-promise-and-pitfalls-of-automatic-content-analysis-methods-for-political-texts/F7AAC8B2909441603FEB25C156448F62",
     "tokenisation, TF-IDF, stopwords, PDF extraction, corpus analysis"),

    (14, 4, "M5: Advanced Python", "Reproducible Research Workflows",
     "Project structure (cookiecutter data science); virtual environments (venv, conda); requirements.txt; git basics for researchers; Jupyter to script conversion; Papermill for parameterised notebooks; documenting code",
     "Reproducibility is now a condition of publication in top journals. A project where you cannot re-run your analysis six months later is a liability. Build these habits now: separate data/code/output folders, never modify raw data, commit changes with git. This day has high long-term ROI even though it feels administrative.",
     "Cookiecutter Data Science project structure | https://drivendata.github.io/cookiecutter-data-science/ | Git for Scientists | https://swcarpentry.github.io/git-novice/",
     "The Turing Way: Guide to Reproducible Research | https://the-turing-way.netlify.app/ | Good Enough Practices in Scientific Computing | https://journals.plos.org/ploscompbiol/article?id=10.1371/journal.pcbi.1005510",
     "None — apply to your Day 12 capstone project by reorganising it into a proper project structure.",
     "Reorganise your Day 12 capstone into a proper directory structure: data/raw/, data/processed/, notebooks/, src/, output/figures/. Create a requirements.txt. Write a README.md explaining the project, data sources, and how to reproduce the analysis.",
     "Read: Wilson, G. et al. (2017). 'Good Enough Practices in Scientific Computing.' PLOS Computational Biology. | https://journals.plos.org/ploscompbiol/article?id=10.1371/journal.pcbi.1005510",
     "project structure, git, virtual environments, reproducibility, documentation"),
]

R_DATA = [
    (1, 1, "M1: R Foundations", "R & RStudio Setup + R Basics",
     "Install R and RStudio; RStudio panes (Console, Script, Environment, Files); vectors; data types (numeric, character, logical, factor, integer); arithmetic; assignment (<-); basic functions: c(), length(), class(), str()",
     "R was built by statisticians for statisticians — its defaults are designed for data analysis, not general programming. The arrow <- for assignment is idiomatic R; use it consistently. Factors are R's way of encoding categorical variables — this matters enormously for regression (R converts them to dummy variables automatically). str() is your first diagnostic on any object.",
     "R for Data Science (2e) Ch.1 | https://r4ds.hadley.nz/ | Swirl package (interactive R tutorial in console) | https://swirlstats.com/",
     "RStudio Education: Beginners | https://education.rstudio.com/learn/beginner/ | Corey Schafer R basics | YouTube | R Programming Tutorial (freeCodeCamp) | https://youtu.be/_V8eKsto3Ug",
     "Use built-in datasets. In console: data() to see available. Use 'airquality', 'mtcars', 'iris'. No download needed.",
     "Create a vector of India's state-level literacy rates (use 5 states from Census 2011). Compute: mean, sd, median, range. Use which() to find states above the national average. Create a factor variable of region (North/South/East/West) and use table() to count how many states per region.",
     "None",
     "vector, factor, numeric, character, logical, str(), class(), assignment"),

    (2, 1, "M1: R Foundations", "Data Structures in R",
     "Matrices; lists; data frames; accessing elements: [], [[]], $; adding/removing rows and columns; rbind(), cbind(); is.na(), which(); basic string functions: paste(), paste0(), nchar(), toupper(), tolower(), gsub()",
     "The data frame is R's version of a spreadsheet — the workhorse of data analysis. The list is R's most flexible structure: it can hold any mix of objects (data frames, vectors, models). Understanding [[]] vs $ for list/df element access prevents constant frustration. is.na() is essential — you will use it every single session.",
     "R for Data Science Ch.20 (Vectors) | https://r4ds.hadley.nz/vectors.html | Advanced R Ch.3 (free) | https://adv-r.hadley.nz/vectors-chap.html",
     "Jenny Bryan's STAT545: Data wrangling | https://stat545.com/ | Hands-On Programming with R (free) | https://rstudio-education.github.io/hopr/",
     "Dataset: Use mtcars (built-in). Explore its structure, extract subsets, add a new column.",
     "Using mtcars: (1) Extract only rows where mpg > 20 using logical indexing; (2) Add a column 'efficiency' = mpg/wt; (3) Create a list containing: the full data frame, the efficiency vector, and a character vector of car names; (4) Access each element of your list using [[]] and $.",
     "None",
     "data frame, list, matrix, indexing, rbind, cbind, is.na, paste, gsub"),

    (3, 1, "M1: R Foundations", "The Tidyverse Part 1 — dplyr",
     "Installing and loading packages; tidyverse philosophy; pipe operator (%>% and |>); filter(); select(); mutate(); arrange(); rename(); distinct(); count(); slice_max(), slice_min()",
     "The tidyverse is a coherent ecosystem of packages that makes R readable and powerful for data work. The pipe operator (%>%) passes the result of one function to the next — it lets you write data transformations as a narrative: 'take the data, then filter, then group, then summarise.' Master the pipe and dplyr and you can handle 80% of all data manipulation tasks.",
     "R for Data Science Ch.3-5 (dplyr) | https://r4ds.hadley.nz/data-transform.html",
     "dplyr cheatsheet (download) | https://github.com/rstudio/cheatsheets/blob/main/data-transformation.pdf | dplyr vignette | https://dplyr.tidyverse.org/articles/dplyr.html | TidyTuesday screencasts by David Robinson | https://www.youtube.com/user/safe4democracy",
     "Dataset: gapminder R package | install.packages('gapminder'). Global life expectancy, GDP, population 1952-2007 across countries. Perfect for development economics exercises.",
     "Using gapminder: (1) Filter to only Asian countries in 2007; (2) Add GDP_total column = gdpPercap * pop; (3) Find the 5 countries with highest life expectancy in 1952 vs 2007; (4) Count countries per continent; (5) Find countries where life expectancy DECREASED between 1952 and 2007 (hint: use arrange and filter after grouping).",
     "None",
     "pipe, filter, select, mutate, arrange, rename, count, slice_max"),

    (4, 2, "M1: R Foundations", "The Tidyverse Part 2 — dplyr Grouped Operations + tidyr",
     "group_by() + summarise(); across(); case_when(); left_join(), right_join(), inner_join(), anti_join(); tidyr: pivot_longer(), pivot_wider(); separate(), unite(); complete(), fill()",
     "Grouped operations are the core of comparative social science analysis. group_by(state) %>% summarise(mean_income = mean(income)) gives you state-by-state averages in two lines. Joins are critical for combining datasets (e.g., survey microdata + administrative boundaries + environmental data). pivot_longer/wider converts between wide (one row per unit) and long (one row per unit-time) formats — you'll need this for panel data and plotting.",
     "R for Data Science Ch.19-20 (joins) + Ch.5 (tidyr) | https://r4ds.hadley.nz/joins.html",
     "tidyr cheatsheet | https://github.com/rstudio/cheatsheets/blob/main/tidyr.pdf | Two-table verbs vignette | https://dplyr.tidyverse.org/articles/two-table.html",
     "Dataset: World Development Indicators via WDI package | install.packages('WDI') | WDI(country='all', indicator=c('SP.POP.TOTL','NY.GDP.PCAP.CD','SE.ADT.LITR.ZS'), start=2000, end=2022) — real panel data.",
     "Using WDI data: (1) Pivot from wide to long format; (2) Compute mean literacy by region (add region via join with a country-region lookup); (3) Use case_when() to create income group categories; (4) Compute each country's GDP growth rate between 2000 and 2022; (5) Use anti_join to find countries in the literacy data but not in the GDP data.",
     "None",
     "group_by, summarise, across, case_when, joins, pivot_longer, pivot_wider"),

    (5, 2, "M2: Visualisation", "ggplot2 Part 1 — Grammar of Graphics",
     "ggplot() + aes() + geom layers; geom_point, geom_line, geom_bar, geom_col, geom_histogram, geom_boxplot, geom_density; colour, fill, size, alpha aesthetics; adding layers; scale functions; coord_flip()",
     "ggplot2 is built on the 'Grammar of Graphics' — every chart is a mapping of data variables to visual properties (x position, y position, colour, size, shape). This makes it systematic: once you understand the grammar, you can construct any chart by specifying the mapping. The key insight: you are not choosing a chart type, you are describing how data maps to visual space.",
     "R for Data Science Ch.2 (ggplot2) | https://r4ds.hadley.nz/data-visualize.html | ggplot2 book (free) | https://ggplot2-book.org/",
     "ggplot2 cheatsheet | https://github.com/rstudio/cheatsheets/blob/main/data-visualization.pdf | R Graph Gallery | https://r-graph-gallery.com/ | TidyTuesday by David Robinson (watch how he builds plots live)",
     "Dataset: gapminder (from Day 3). Replicate the famous Hans Rosling bubble chart: GDP per capita (x) vs life expectancy (y), bubble size = population, colour = continent, for year 2007.",
     "Using gapminder 2007: (1) Reproduce the Rosling bubble chart; (2) Create a line chart showing life expectancy over time for India, China, USA, Brazil; (3) Create side-by-side boxplots of life expectancy by continent; (4) Create a histogram of GDP per capita with a density overlay. Add proper titles, axis labels, and caption with data source to all plots.",
     "None",
     "ggplot, aes, geom layers, scales, aesthetics, grammar of graphics"),

    (6, 2, "M2: Visualisation", "ggplot2 Part 2 — Publication-Quality Charts",
     "facet_wrap() and facet_grid(); themes (theme_minimal, theme_bw, theme_classic); custom theme() modifications; scale_color_brewer, scale_fill_viridis; labels with geom_text and ggrepel; annotation; saving with ggsave(); patchwork for multi-panel",
     "Academic publications require specific formatting: no grey background, clear fonts, colourblind-safe palettes, proper margins. The viridis palette is colourblind-safe and prints well in greyscale — use it by default. patchwork lets you combine multiple ggplots into one figure (A, B, C panels). Learning to make journal-ready figures now means you won't need to redo them later.",
     "ggplot2 book Ch.17-18 (Themes) | https://ggplot2-book.org/polishing | patchwork package | https://patchwork.data-imaginist.com/",
     "ggthemes package | https://github.com/jrnold/ggthemes | ggrepel | https://github.com/slowkow/ggrepel | Cedric Scherer's ggplot2 workshop | https://www.cedricscherer.com/2019/08/05/a-ggplot2-tutorial-for-beautiful-plotting-in-r/",
     "Dataset: India COVID vaccination data (public) | https://api.covid19india.org/ or download from https://www.kaggle.com/datasets/sudalairajkumar/covid19-in-india",
     "Create a 4-panel publication figure (using patchwork): (A) line chart of vaccinations over time for 5 states; (B) bar chart of total doses by state; (C) scatter plot vaccinations vs population; (D) summary table as a ggplot (using ggtext or gt package). Apply theme_minimal, viridis palette, ggrepel labels. Save at 300 dpi.",
     "None",
     "facet_wrap, themes, viridis, patchwork, ggsave, ggrepel, annotation"),

    (7, 3, "M3: R for Research", "R Markdown + Quarto: Reproducible Documents",
     "R Markdown syntax; code chunks and options (echo, eval, warning, message); inline R code; YAML header; output formats (HTML, PDF via LaTeX, Word); cross-references; bibliographies with BibTeX; Quarto basics",
     "R Markdown is how professional quantitative researchers in social science write papers and reports. You write text and code in the same document — when you 'knit' it, R runs all the code and inserts outputs (tables, charts) automatically. This means your figures and numbers in the text are always in sync with your analysis. Many journals now accept R Markdown manuscripts directly.",
     "R Markdown: The Definitive Guide (free) | https://bookdown.org/yihui/rmarkdown/ | Quarto (next-gen R Markdown) | https://quarto.org/docs/get-started/",
     "RStudio: Introduction to R Markdown | https://rmarkdown.rstudio.com/lesson-1.html | Alison Hill's R Markdown workshop | https://alison.rbind.io/project/rmd4medicine/",
     "Take your Day 12 Python capstone data/results and reproduce them in R Markdown instead.",
     "Write a 2-page R Markdown report on any India development dataset: abstract (written in markdown), data section (described with inline R), analysis section (2 regressions with stargazer tables), results section (2 ggplot figures), and conclusion. Knit to both HTML and PDF. This is the format of a thesis chapter.",
     "None",
     "R Markdown, YAML, chunk options, knitr, BibTeX, inline code, Quarto"),

    (8, 3, "M3: R for Research", "Statistical Tests + Modelling Foundations in R",
     "t.test(), wilcox.test(), chisq.test(), cor.test(), aov(); lm() basics: formula syntax, summary.lm(), residuals, fitted values; coefplot; confidence intervals with confint(); broom package for tidy model output",
     "R's formula syntax (outcome ~ predictor1 + predictor2) is elegant and consistent across all modelling functions — learn it once and it works for lm, glm, lmer, spatialreg. The broom package converts messy model output into tidy data frames, making it easy to plot coefficients or compare models in a table. Always examine residuals — a plot of residuals vs fitted values reveals model misspecification.",
     "R for Data Science Ch.24-26 (modelling) | https://r4ds.hadley.nz/ | broom package | https://broom.tidymodels.org/",
     "ModernDive: Statistical Inference via Data Science (free) | https://moderndive.com/ | StatQuest R playlist | https://www.youtube.com/playlist?list=PLblh5JKOoLUKAtDViTvRGFpphEc24M-QH",
     "Dataset: NFHS-5 district summary | https://rchiips.org/nfhs/districtfactsheet_NFHS-5.shtml — download PDF tables or use pre-cleaned version from https://github.com/PerformingStatistics/NFHS5",
     "Using NFHS district data: (1) t-test: is stunting significantly different between EAG and non-EAG states?; (2) chi-square: is there an association between toilet access category and stunting severity category?; (3) OLS regression: predict under-5 mortality from anaemia, toilet access, and mother's education; (4) Report all results using broom::tidy() and display with kable().",
     "None",
     "t.test, chisq.test, lm, formula syntax, broom, confint, residual diagnostics"),

    (9, 3, "M3: R for Research", "Advanced dplyr + Data Programming",
     "Programming with dplyr: {{ }} embracing, .data pronoun, across() with purrr-style functions; purrr: map(), map_df(), map2(); functional programming concepts; writing reusable analysis functions; rowwise operations",
     "Once you move beyond exploratory analysis to writing reproducible research pipelines, you need to write functions that work inside dplyr pipelines. purrr's map() replaces for loops for applying functions across lists (e.g., run the same regression model on 30 different countries and collect results). This is the difference between 'I can analyse one dataset' and 'I can run systematic comparative analyses.'",
     "R for Data Science Ch.26 (Functions) + Ch.27 (Iteration) | https://r4ds.hadley.nz/ | purrr tutorial | https://jennybc.github.io/purrr-tutorial/",
     "Advanced R: Functional Programming | https://adv-r.hadley.nz/fp.html | Rebecca Barter's purrr tutorial | https://www.rebeccabarter.com/blog/2019-08-19_purrr",
     "Dataset: gapminder — run country-level regressions for all Asian countries.",
     "For each Asian country in gapminder: (1) Write a function that runs lm(lifeExp ~ year) and returns slope, R², and p-value as a named list; (2) Use purrr::map_df() to apply this to all Asian countries; (3) Visualise the distribution of slopes; (4) Identify the 5 countries with the fastest and slowest improvement in life expectancy. Write 150-word interpretation.",
     "Read: Grolemund & Wickham. 'R for Data Science' — iteration chapter. Also read the purrr cheatsheet carefully.",
     "{{ }}, .data, across, purrr::map, map_df, functional programming, iteration"),

    (10, 4, "M4: R Capstone", "R Capstone: Replication Exercise",
     "Full replication of a published paper's quantitative analysis; reading academic R code; understanding variable construction; interpreting and reproducing tables and figures; writing a replication note",
     "The highest-leverage learning exercise is to take a published paper with available data and code and replicate it exactly. This teaches you how professional researchers structure their code, how they construct variables, and how their reported results map to their analysis. Finding discrepancies teaches you more than any tutorial.",
     "Harvard Dataverse (replication archives) | https://dataverse.harvard.edu/ | ICPSR replication data | https://www.icpsr.umich.edu/",
     "Open Science Framework | https://osf.io/ | Political Science Replication | https://github.com/erikgahner/PolData",
     "Choose any replicated paper from Harvard Dataverse with India data. Suggested: Banerjee & Duflo (2007) 'The Economic Lives of the Poor' data via ICPSR | https://www.icpsr.umich.edu/web/ICPSR/studies/22626",
     "Download data and code from your chosen paper. Run the code. Reproduce the main table and one figure. Document any discrepancies. Write a 300-word replication note describing what you did and any issues you encountered. This is real research practice.",
     "Read: Clemens, M.A. (2017). 'The meaning of failed replications.' Journal of Economic Methodology. | https://doi.org/10.1080/1350178X.2017.1366069",
     "Replication, code reading, variable construction, research integrity"),
]

STATS_DATA = [
    (1, 1, "M1: Descriptive Statistics", "Distributions and Data Summarisation",
     "Types of variables (nominal, ordinal, interval, ratio); central tendency (mean, median, mode); spread (variance, SD, IQR, range); skewness and kurtosis; frequency tables; empirical vs theoretical distributions; QQ plots",
     "Before any modelling: describe your data exhaustively. What is the shape of the distribution? Is it skewed (most income data is right-skewed)? Are there outliers? What is the unit of observation? These questions determine every modelling choice downstream. For social science data: always report medians alongside means for skewed distributions. Always check if your continuous variables have meaningful zeros (ratio vs interval).",
     "Statistics (OpenStax, free) Ch.1-2 | https://openstax.org/books/statistics/pages/1-introduction | Descriptive Statistics (Khan Academy) | https://www.khanacademy.org/math/statistics-probability/summarizing-quantitative-data",
     "Naked Statistics by Wheelan (read as prose) | StatQuest: 'Statistics Fundamentals' | https://youtube.com/playlist?list=PLblh5JKOoLUK0FLuzwntyYI10UQFUhsY9 | The Art of Statistics by Spiegelhalter",
     "Dataset: India Human Development Survey (IHDS) | https://ihds.umd.edu/ihds-data — download Wave 2 (2011-12). Contains income, education, caste, health, asset data for 42,000 households.",
     "Using IHDS: (1) Compute mean, median, mode for household income — why do they differ so much?; (2) Compute IQR and identify outliers by IQR rule; (3) Plot income distribution (histogram + QQ plot); (4) Compute skewness coefficient (e_1071 package in R or scipy.stats in Python); (5) Transform income using log and re-examine distribution. Write 200-word commentary.",
     "Read: Anscombe, F.J. (1973). 'Graphs in Statistical Analysis.' American Statistician, 27(1). — The Anscombe Quartet: why always visualise before computing statistics.",
     "nominal/ordinal/interval/ratio, mean/median/mode, IQR, skewness, QQ plot, log transformation"),

    (2, 1, "M1: Descriptive Statistics", "Probability and Sampling Theory",
     "Probability rules; conditional probability; Bayes' theorem (intuition); Law of Large Numbers; Central Limit Theorem (CLT); standard error; sampling distributions; types of sampling (SRS, stratified, cluster, multi-stage); sampling weights",
     "The Central Limit Theorem is the reason all parametric statistics work: sample means are normally distributed even when the underlying variable is not, provided sample size is large enough. This is why t-tests work on income data. Sampling weights are critical for India survey data — NFHS and IHDS both use complex multi-stage sampling, and ignoring weights produces biased estimates at the national level.",
     "OpenStax Statistics Ch.6-8 | https://openstax.org/books/statistics/ | Khan Academy Probability | https://www.khanacademy.org/math/statistics-probability/probability-library",
     "StatQuest: 'The Central Limit Theorem' | https://youtu.be/YAlJCEDH2uY | StatQuest: 'Sampling distributions' | All of Statistics Ch.3-4 (Wasserman, free) | https://link.springer.com/book/10.1007/978-0-387-21736-9",
     "Simulate the CLT: draw samples of varying sizes from a highly skewed distribution (use R/Python) and show convergence of sample mean distribution to normal. Use NFHS complex sample weights in R with survey package.",
     "In R using the survey package: (1) Load NFHS district data; (2) compute weighted vs unweighted national mean of under-5 mortality; (3) compute standard error using complex sample design; (4) construct 95% CI; (5) Simulate CLT: draw 1000 samples of size 30 from a right-skewed income distribution, plot the sampling distribution of the mean. Show it is approximately normal.",
     "Read: Kish, L. (1995). 'The 150-year inheritance.' Invited address — on the history and practice of survey sampling. OR Read survey package vignette | https://cran.r-project.org/web/packages/survey/vignettes/survey.pdf",
     "CLT, standard error, sampling distribution, sampling weights, complex survey design"),

    (3, 2, "M2: Inference", "Hypothesis Testing",
     "Null and alternative hypotheses; Type I and Type II errors; p-values (what they mean and what they don't); statistical vs practical significance; effect sizes (Cohen's d, eta-squared, Cramér's V); one-tail vs two-tail tests; multiple testing problem; Bonferroni correction",
     "The p-value crisis in social science: p < 0.05 does not mean the effect is real, important, or replicable. It only means the data are unlikely under the null hypothesis. Effect sizes tell you if the effect matters practically — a statistically significant reduction in stunting rates of 0.001 percentage points is not a policy-relevant finding. Always report both. The multiple testing problem (running 20 tests and finding 1 significant result) is the source of much false discovery in empirical social science.",
     "OpenStax Statistics Ch.9-10 | https://openstax.org/books/statistics/ | Khan Academy: Significance Tests | https://www.khanacademy.org/math/statistics-probability/significance-tests-one-sample",
     "StatQuest: 'P-values' | https://youtu.be/vemZtEM63GY | ASA Statement on P-values | https://www.tandfonline.com/doi/abs/10.1080/00031305.2016.1154108 | Wasserstein & Lazar (2016) — read this carefully",
     "Dataset: Replicate (or critically examine) results from any paper that uses p<0.05 as a threshold. Try: Duflo, Dupas, Kremer (2011) on education in Kenya (data via AEA) | https://www.aeaweb.org/articles?id=10.1257/aer.101.5.1739",
     "Using NFHS state data: (1) For each of 15 health indicators, test if there is a significant difference between EAG and non-EAG states; (2) Apply Bonferroni correction — how many remain significant?; (3) For each significant result, compute Cohen's d; (4) Report in a table showing: indicator, t-stat, p-value (raw and Bonferroni-corrected), Cohen's d. Discuss implications for how you would report findings.",
     "Read: Ioannidis, J.P.A. (2005). 'Why Most Published Research Findings Are False.' PLOS Medicine. | https://doi.org/10.1371/journal.pmed.0020124 — one of the most cited papers in science.",
     "null hypothesis, p-value, Type I/II error, effect size, Cohen's d, multiple testing, Bonferroni"),

    (4, 2, "M2: Inference", "Regression I — Simple and Multiple OLS",
     "OLS derivation (intuition, not calculus); interpretation of slope and intercept; R-squared and adjusted R-squared; standard errors and t-statistics; confidence intervals for coefficients; categorical predictors and dummy variables; interaction terms; reading a regression table",
     "Regression is the foundational tool of empirical social science. The coefficient on a variable tells you the expected change in the outcome for a one-unit change in the predictor, holding all other predictors constant — the 'ceteris paribus' clause. Interaction terms (X1*X2) are essential for testing whether an effect varies by group (e.g., does the effect of education on income differ by caste?). Always interpret coefficients substantively, not just statistically.",
     "Stock & Watson 'Introduction to Econometrics' Ch.4-7 (the clearest regression textbook) | Wooldridge 'Introductory Econometrics' Ch.2-4",
     "Econometrics by Simulation (blog + YouTube) | https://www.econometricsbysimulation.com/ | StatQuest: Linear Regression | https://youtu.be/7ArmBVF2dCs | ISLR Ch.3 (free) | https://www.statlearning.com/",
     "Dataset: India Human Development Survey (IHDS) | https://ihds.umd.edu/ — regress log household income on education, land ownership, caste, state dummies.",
     "Using IHDS: (1) Regress log(income) on years of education alone — interpret slope; (2) Add controls: land ownership, caste (as factor), state (fixed effects); interpret each coefficient; (3) Add interaction: education × SC/ST dummy — does education's effect differ by caste?; (4) Report using stargazer (R) or statsmodels.summary_col (Python) — reproduce as a proper regression table with column headers and N, R² at the bottom.",
     "Read: Angrist & Pischke (2010). 'The Credibility Revolution in Empirical Economics.' Journal of Economic Perspectives, 24(2). | https://www.aeaweb.org/articles?id=10.1257/jep.24.2.3",
     "OLS, slope interpretation, dummy variables, interactions, adjusted R², regression table, stargazer"),

    (5, 2, "M2: Inference", "Regression II — Diagnostics and Robust SE",
     "OLS assumptions (LINE: Linearity, Independence, Normality of residuals, Equal variance/homoskedasticity); residual plots; heteroskedasticity (Breusch-Pagan test); multicollinearity (VIF); robust standard errors (HC1-HC3); clustered standard errors; log transformations; outlier analysis (Cook's D, leverage)",
     "OLS assumptions are almost always violated in social science data. Heteroskedasticity (variance of errors varies with X — extremely common in cross-sectional income data) biases standard errors. The fix: robust standard errors. Clustered standard errors are needed when observations are not independent within groups (e.g., students within schools, districts within states). VIF > 10 indicates serious multicollinearity. Learn to diagnose first, then correct.",
     "Wooldridge Ch.8 (Heteroskedasticity) | https://www.cengage.com/c/introductory-econometrics-a-modern-approach-7e-wooldridge/ | Stock & Watson Ch.5 (Robust SE)",
     "StatQuest: 'Residual Plots' | https://youtu.be/vpk_1gldOAE | Sandwich package documentation (R) | https://sandwich.r-forge.r-project.org/ | HC robust SE in statsmodels (Python) | https://www.statsmodels.org/stable/generated/statsmodels.regression.linear_model.OLSResults.get_robustcov_results.html",
     "Dataset: Repeat your Day 4 IHDS regression. Run full diagnostics.",
     "On your IHDS regression: (1) Plot residuals vs fitted — is there a pattern?; (2) Run Breusch-Pagan test for heteroskedasticity; (3) Compare standard errors: OLS vs HC2 robust vs clustered (by state); (4) Compute VIF for all predictors — any multicollinearity?; (5) Identify the 5 observations with highest Cook's D — are they influential? Re-run without them and compare. Write 200-word 'diagnostic summary' as would appear in a paper appendix.",
     "Read: MacKinnon & White (1985) on heteroskedasticity-robust standard errors. OR more accessible: Hayes & Cai (2007). 'Using heteroskedasticity-consistent standard error estimators.' Behavior Research Methods.",
     "LINE assumptions, Breusch-Pagan, VIF, robust SE, clustered SE, Cook's D, leverage"),

    (6, 3, "M3: Limited Dependent Variables", "Logistic Regression and Probit",
     "Binary outcomes; why OLS fails for binary outcomes (linear probability model problems); logistic function; maximum likelihood estimation (intuition); logit and probit; odds ratios; average marginal effects; goodness of fit (McFadden R², Hosmer-Lemeshow); ROC curves",
     "Many outcomes in development research are binary: has toilet/does not, stunted/not stunted, migrated/did not. OLS on binary outcomes (linear probability model) predicts probabilities outside [0,1] and has inherently heteroskedastic errors. Logistic regression constrains predictions to [0,1] via the logistic function. The key interpretation challenge: logit coefficients are in log-odds — always convert to average marginal effects (AME) for substantive interpretation.",
     "Wooldridge Ch.17 (limited dependent variables) | UCLA IDRE Logistic Regression | https://stats.oarc.ucla.edu/r/dae/logit-regression/",
     "StatQuest: 'Logistic Regression' | https://youtu.be/vN5cNN2-HWE | margins package in R | https://cran.r-project.org/web/packages/margins/ | marginaleffects package (recommended) | https://marginaleffects.com/",
     "Dataset: NFHS-5 individual data on child stunting (binary) | Download from DHS Program | https://dhsprogram.com/data/dataset/India_Standard-DHS_2019.cfm — requires free registration.",
     "Using NFHS child data: (1) Estimate LPM (OLS) and logit for stunting ~ mother_education + toilet + wealth_index + child_age + sex; (2) Compare coefficient signs — do they agree?; (3) For logit: compute AME using marginaleffects package; (4) Interpret: what is the predicted probability of stunting for a child with educated vs uneducated mother, holding all else constant?; (5) Plot ROC curve, compute AUC.",
     "Read: Mood, C. (2010). 'Logistic Regression: Why We Cannot Do What We Think We Can Do.' European Sociological Review, 26(1). | https://academic.oup.com/esr/article/26/1/67/579890",
     "binary outcomes, logistic function, MLE, odds ratios, average marginal effects, McFadden R²"),

    (7, 3, "M3: Limited Dependent Variables", "Count Data and Ordered Models",
     "Poisson regression; negative binomial regression (overdispersion); zero-inflated models; ordered logit and probit; multinomial logit; IIA assumption; testing for overdispersion",
     "Count outcomes (number of hospital visits, number of children, number of days of work missed) require Poisson or negative binomial regression — not OLS. In practice, count data is almost always overdispersed (variance > mean), so negative binomial is the default over Poisson. Ordered outcomes (education level 1-5, wealth quintile) need ordered logit — not OLS on the rank number, which imposes equal spacing between categories.",
     "Wooldridge Ch.17 | Cameron & Trivedi 'Microeconometrics Using Stata' Ch.18-20 | Long 'Regression Models for Categorical and Limited Dependent Variables'",
     "UCLA IDRE: Count Data Models | https://stats.oarc.ucla.edu/r/dae/negative-binomial-regression/ | MASS package in R | DataCamp: Count Data tutorial",
     "Dataset: IHDS-II — number of doctor visits in past year (count variable). Also: NFHS birth history data for children ever born (count).",
     "Using IHDS: (1) Regress number of doctor visits on income, age, caste, state; run OLS first, identify why it is inappropriate (negative predictions?); (2) Run Poisson — test for overdispersion (dispersion test in AER package); (3) Run negative binomial; (4) Compare predicted vs actual distribution from each model; (5) Compute incidence rate ratios (IRR) from negative binomial and interpret.",
     "Read: Winkelmann, R. (2008). 'Econometric Analysis of Count Data.' Ch.1-3 for conceptual grounding.",
     "Poisson, negative binomial, overdispersion, zero-inflation, ordered logit, IRR"),

    (8, 4, "M4: Time Series + Panel Data Basics", "Time Series Fundamentals",
     "Stationarity and unit roots (ADF test); autocorrelation (ACF/PACF plots); ARIMA: AR, I, MA components; forecasting; spurious regression; cointegration concept; HAC standard errors",
     "Spurious regression is the classic trap: two non-stationary time series can show high R² and significant coefficients even with no real relationship (e.g., Indian GDP and US cheese production are both trending up, so they appear correlated). Always test for unit roots before regressing one time series on another. For most social science applications you will use ARIMA for description/forecasting and first-differencing or detrending to achieve stationarity.",
     "Hyndman & Athanasopoulos 'Forecasting: Principles and Practice' (free) | https://otexts.com/fpp3/ | Stock & Watson Ch.14-16",
     "forecast package in R | https://pkg.robjhyndman.com/forecast/ | tseries package | StatQuest: 'ARIMA' | https://youtu.be/Aw77aMLj9uM",
     "Dataset: RBI Handbook of Statistics on Indian Economy | https://rbi.org.in/Scripts/AnnualPublications.aspx?head=Handbook%20of%20Statistics%20on%20Indian%20Economy — GDP, inflation, employment quarterly series.",
     "Using RBI time series: (1) Plot India GDP quarterly series; (2) Run ADF test — is it stationary?; (3) Take first differences — re-test; (4) Plot ACF and PACF of differenced series; (5) Fit ARIMA(p,d,q) — use auto.arima() to select order; (6) Forecast 4 quarters ahead with 95% PI; (7) Plot actual vs fitted with uncertainty bands.",
     "Read: Nelson & Plosser (1982). 'Trends and Random Walks in Macroeonomic Time Series.' Journal of Monetary Economics — the paper that established unit roots matter.",
     "stationarity, unit root, ADF test, ACF/PACF, ARIMA, spurious regression, HAC SE"),

    (9, 4, "M4: Time Series + Panel Data Basics", "Index Construction and Composite Indicators",
     "Principal Component Analysis (PCA) for dimensionality reduction; factor analysis; min-max normalisation; z-score normalisation; weighting schemes (equal, PCA-derived, expert); Alkire-Foster method; robustness checks (sensitivity analysis, Pearson's correlation between index variants)",
     "You have direct experience with the AF method from your MPI proposal. This session systematises index construction and teaches the leading alternatives. PCA is the most common data-driven weighting method — it derives weights from the variance structure of the indicators. Key principle: always run a sensitivity analysis showing how your index ranking changes with different weight schemes and normalisation methods. Your index is credible only if rankings are stable.",
     "OECD Handbook on Composite Indicators (free) | https://www.oecd.org/en/publications/handbook-on-constructing-composite-indicators_9789264043466-en.html | OPHI AF Method documentation | https://ophi.org.uk/research/multidimensional-poverty/mpi-methodology/",
     "FactoMineR package for PCA in R | https://cran.r-project.org/web/packages/FactoMineR/ | factoextra for PCA visualisation | PCA intuition: StatQuest | https://youtu.be/FgakZw6K1QQ",
     "Dataset: India NFHS-5 district data (15+ health/development indicators) | https://rchiips.org/nfhs/ — construct a multi-dimensional development index.",
     "Using NFHS district data with 10 indicators: (1) Run PCA — how many components explain 80% of variance?; (2) Construct an equal-weight index and a PCA-weighted index; (3) Compute rank correlation between the two — do they agree?; (4) Identify 5 districts that rank very differently under the two methods — why?; (5) Construct an AF-style counting index using 3 of the indicators with a k=2 cutoff. Map all three index variants.",
     "Read: Alkire, S. & Foster, J. (2011). 'Counting and Multidimensional Poverty Measurement.' Journal of Public Economics, 95(7-8), 476-487. | https://www.sciencedirect.com/science/article/pii/S0047272710001666",
     "PCA, factor analysis, normalisation, weighting, AF method, sensitivity analysis, index robustness"),

    (10, 5, "M5: Bayesian Basics", "Introduction to Bayesian Thinking for Social Scientists",
     "Frequentist vs Bayesian framework; prior, likelihood, posterior; Bayes' theorem; credible intervals vs confidence intervals; MCMC intuition (no maths); rstanarm/brms packages in R; when to use Bayesian methods; interpreting posterior distributions",
     "Bayesian analysis is increasingly expected in social science, particularly for small samples, complex hierarchical models, and when you want to formally incorporate prior knowledge. The conceptual shift: in Bayesian analysis, you treat parameters as random variables with probability distributions rather than as fixed unknowns. A 95% credible interval means 'there is a 95% probability the true parameter lies here' — which is what most researchers incorrectly believe a frequentist CI means.",
     "Statistical Rethinking by Richard McElreath (the best Bayesian textbook for social scientists) | https://xcelab.net/rm/ | brms package | https://paul-buerkner.github.io/brms/",
     "Statistical Rethinking 2023 lectures (free, YouTube) | https://youtube.com/playlist?list=PLDcUM9US4XdPz-KxHM4XHt7uUVGWWVHUs | Bayes Factors tutorial | https://easystats.github.io/bayestestR/",
     "Dataset: Use gapminder data. Run a Bayesian multilevel model of life expectancy across countries using brms.",
     "Using rstanarm: (1) Fit a Bayesian linear regression predicting life expectancy from GDP — compare posterior distribution of slope to frequentist OLS confidence interval; (2) Interpret the credible interval; (3) Change the prior on the slope from flat to weakly informative — does it change results?; (4) Plot the posterior distribution using bayesplot package. Write 200-word conceptual explanation of what the posterior tells you.",
     "Read: Gelman, A. et al. (2020). 'Bayesian Workflow.' ArXiv preprint. | https://arxiv.org/abs/2011.01808 — a pragmatic guide to using Bayesian methods.",
     "prior, posterior, likelihood, credible interval, MCMC, rstanarm, brms, posterior distribution"),
]

ECON_DATA = [
    (1, 1, "M1: OLS Revisited", "Gauss-Markov Theorem and OLS Assumptions",
     "BLUE (Best Linear Unbiased Estimator) properties; OLS assumptions formally stated; what each assumption means in social science practice; what happens when each is violated; omitted variable bias (OVB) — direction and magnitude; algebraic OVB formula",
     "OVB is the central problem of observational social science. If you omit a variable that (a) is correlated with your treatment variable and (b) affects the outcome, your estimated treatment effect is biased. Crucially, you can determine the direction of the bias using the OVB formula even without the omitted variable. This is the reasoning behind causal diagrams. Understand this deeply — it motivates every identification strategy in the remaining days.",
     "Wooldridge 'Introductory Econometrics' Ch.3-4 (omitted variable bias) | Angrist & Pischke 'Mostly Harmless Econometrics' Ch.3",
     "Mostly Harmless Econometrics (free PDF via ResearchGate) | https://www.researchgate.net/publication/51992844 | Nick Huntington-Klein 'The Effect' (free) | https://theeffectbook.net/ | YouTube: Ben Lambert Econometrics | https://www.youtube.com/channel/UC3tFZR3eL1bDY8CqZDOQh-w",
     "Dataset: Mincer earnings equation — IHDS data. Estimate returns to education with and without ability proxy (use household asset index).",
     "Using IHDS: (1) Regress log income on education — call this β1; (2) Add father's education as ability proxy — call new coefficient β2; (3) Manually compute the OVB: β1 - β2 = (coefficient of father_edu in step 2) × (coefficient of father_edu in auxiliary regression of education on father_edu); (4) Verify your formula gives the right answer; (5) Write a paragraph explaining the direction and source of bias in the naive regression.",
     "Read: Angrist, J. & Pischke, J-S. (2010). 'The Credibility Revolution in Empirical Economics.' JEP, 24(2). | https://www.aeaweb.org/articles?id=10.1257/jep.24.2.3",
     "BLUE, OVB formula, direction of bias, omitted confounders, Gauss-Markov"),

    (2, 1, "M1: OLS Revisited", "Instrumental Variables and 2SLS",
     "Endogeneity: simultaneity, measurement error, OVB; IV assumptions: relevance, exclusion restriction, exogeneity; two-stage least squares (2SLS) mechanics; weak instruments (F-statistic rule of thumb); Hausman endogeneity test; IV in R (AER::ivreg) and Python (linearmodels)",
     "IV is the workhorse of causal identification in observational data. The key is finding an instrument Z that (1) correlates with the treatment X (relevance) but (2) affects the outcome Y ONLY through X (exclusion restriction). The exclusion restriction is untestable — it requires a theoretical argument. Classic instruments: rainfall as an instrument for agricultural income; distance to a reform city as instrument for policy adoption. The exclusion restriction must be defended theoretically, not statistically.",
     "Wooldridge Ch.15 | Angrist & Pischke 'Mostly Harmless Econometrics' Ch.4 | Stock & Watson Ch.12",
     "The Effect Book: IV chapter | https://theeffectbook.net/ch-InstrumentalVariables.html | AER package for ivreg | https://cran.r-project.org/web/packages/AER/ | Nick HK: 'IV Introduction' video",
     "Dataset: Angrist & Krueger (1991) quarter-of-birth as IV for education — data via ICPSR. OR use rainfall-income IV in an India agricultural dataset from ICRISAT | https://vdsa.icrisat.org/vdsa-mospi.aspx",
     "Using a dataset with a valid instrument: (1) Estimate naive OLS of income on education; (2) Argue endogeneity theoretically; (3) Show instrument relevance: first stage regression, report F-statistic; (4) Run 2SLS using ivreg(); (5) Run Hausman test — is endogeneity statistically confirmed?; (6) Compare OLS vs 2SLS estimates and interpret the LATE (Local Average Treatment Effect).",
     "Read: Angrist, J. & Krueger, A. (1991). 'Does Compulsory School Attendance Affect Schooling and Earnings?' QJE. The canonical IV paper. | https://academic.oup.com/qje/article-abstract/106/4/979/1873496",
     "endogeneity, relevance, exclusion restriction, 2SLS, first stage, F-statistic, Hausman test, LATE"),

    (3, 2, "M2: Panel Data", "Fixed and Random Effects",
     "Panel data structure (units × time); within-variation vs between-variation; Fixed Effects (FE): demeaning, entity fixed effects, time fixed effects, two-way FE; Random Effects (RE); Hausman test for FE vs RE; plm package in R; standard errors: clustered by unit",
     "Fixed effects are the fundamental tool for controlling for unobserved time-invariant heterogeneity. If you have panel data on Indian districts across multiple years, a district FE controls for all time-invariant district characteristics (geography, history, culture) that you cannot observe or measure. The Hausman test tells you whether the RE assumption (unobserved effects uncorrelated with X) is valid — in social science it almost never is, so FE is the default. Always cluster SEs at the unit level.",
     "Wooldridge Ch.13-14 | Angrist & Pischke Ch.5 | Croissant & Millo 'Panel Data Econometrics in R' | https://cran.r-project.org/web/packages/plm/vignettes/plmPackage.html",
     "plm package | https://cran.r-project.org/web/packages/plm/ | The Effect Book: Panel Data | https://theeffectbook.net/ch-FixedEffects.html | linearmodels Python | https://bashtage.github.io/linearmodels/",
     "Dataset: World Bank panel — states × years from SHRUG | https://www.devdatalab.org/shrug_download/ — district-level panel 2001-2011 with economic, population, infrastructure data.",
     "Using SHRUG district panel: (1) Run pooled OLS of nighttime lights (economic proxy) on road access; (2) Add district FE (plm, model='within'); (3) Add year FE (two-way FE); (4) Run Hausman test; (5) Cluster SEs at district level; (6) Interpret: does the within-district effect of road access on economic activity differ from the cross-sectional OLS estimate? Why? Write 250-word interpretation.",
     "Read: Bertrand, M., Duflo, E. & Mullainathan, S. (2004). 'How Much Should We Trust Difference-in-Differences Estimates?' QJE, 119(1). — on the importance of clustered SE in panel data.",
     "FE, RE, demeaning, within variation, two-way FE, Hausman test, clustered SE, plm"),

    (4, 2, "M2: Panel Data", "Difference-in-Differences (DiD)",
     "Parallel trends assumption; 2×2 DiD; generalised DiD with controls; event study design; staggered DiD (Callaway-Sant'Anna estimator); placebo tests; pre-trends test; did and did2s packages in R",
     "DiD is the most widely used quasi-experimental method in development economics and public policy evaluation. The logic: compare the change in outcome for a treated group to the change for a control group over the same period. The key assumption: absent treatment, both groups would have followed parallel trends. Testing pre-trends (showing parallel pre-period trends) is now obligatory in top journals. The staggered DiD problem (when units are treated at different times) has fundamentally changed applied practice since 2019.",
     "Angrist & Pischke Ch.5 | Callaway & Sant'Anna (2021) paper | The Effect Book: DiD | https://theeffectbook.net/ch-DifferenceinDifference.html",
     "did package (Callaway-Sant'Anna) | https://bcallaway11.github.io/did/ | mixtape (Scott Cunningham, free) | https://mixtape.scunning.com/09-difference_in_differences | Baker et al. (2022) on staggered DiD",
     "Dataset: India's Pradhan Mantri Gram Sadak Yojana (PMGSY) road construction rollout as treatment. SHRUG data has village-level connectivity + outcomes. OR: MGNREGS rollout across districts.",
     "Using PMGSY rollout data from SHRUG: (1) Define treatment (road connected) and control (not yet connected) groups; (2) Run 2×2 DiD manually first; (3) Run generalised DiD with controls and year FE; (4) Plot event study graph (coefficients on each year relative to treatment); (5) Test pre-trends; (6) Discuss: what violations of parallel trends are most plausible here?",
     "Read: Callaway, B. & Sant'Anna, P. (2021). 'Difference-in-Differences with Multiple Time Periods.' Journal of Econometrics, 225(2). | https://www.sciencedirect.com/science/article/pii/S0304407620303948",
     "parallel trends, 2×2 DiD, event study, pre-trends test, staggered DiD, Callaway-Sant'Anna"),

    (5, 3, "M3: Quasi-Experiments", "Regression Discontinuity Design (RDD)",
     "Sharp vs fuzzy RDD; running variable; bandwidth selection; local linear regression; McCrary density test; continuity assumption; polynomial order choice; rdrobust package; visualisation of RD",
     "RDD exploits the arbitrary rule of a cutoff: units just above and just below a threshold are assumed to be as good as randomly assigned. In social policy: income cutoffs for program eligibility (MGNREGS income threshold), population cutoffs for local governance rules, exam score cutoffs for scholarships. The validity depends on: (1) units cannot precisely manipulate their assignment to be just above/below the cutoff; (2) no other policy changes at the same cutoff.",
     "Lee & Lemieux (2010) 'Regression Discontinuity Designs in Economics.' Journal of Economic Literature. | rdrobust package | https://rdpackages.github.io/rdrobust/",
     "The Effect Book: RDD | https://theeffectbook.net/ch-RegressionDiscontinuity.html | NBER RDD lectures (Imbens & Lemieux) | https://www.nber.org/papers/t0337 | rddtools package",
     "Dataset: MGNREGS targeting — districts above/below poverty threshold. OR: Indian election results data at vote-share cutoff (50%) for close elections. Download: Lok Sabha results from Election Commission | https://eci.gov.in/",
     "Using close election data: (1) Define running variable (vote share margin relative to 50%); (2) Plot outcome (e.g., public goods investment) against running variable with local linear fit on each side of cutoff; (3) Run rdrobust() — report point estimate and 95% CI; (4) Test density continuity at cutoff (McCrary/rddensity test); (5) Try polynomial orders 1, 2, 3 — do results change? Interpret.",
     "Read: Lee, D. & Lemieux, T. (2010). 'Regression Discontinuity Designs in Economics.' Journal of Economic Literature, 48(2), 281-355. | https://www.aeaweb.org/articles?id=10.1257/jel.48.2.281",
     "running variable, sharp/fuzzy RDD, bandwidth, local linear regression, McCrary test, rdrobust"),

    (6, 3, "M3: Quasi-Experiments", "Propensity Score Matching (PSM) and Causal Inference",
     "Potential outcomes framework (Rubin); ATE vs ATT vs ATU; confounding and backdoor paths; DAGs (directed acyclic graphs); propensity score estimation; nearest neighbour matching; kernel matching; covariate balance (SMD); sensitivity analysis (Rosenbaum bounds)",
     "PSM is useful when you have observational data and cannot randomise, but you have a rich enough set of covariates to make the 'selection on observables' assumption credible. The key step is checking covariate balance after matching — if treated and control units look similar on all measured covariates after matching, your estimate is credible. Rosenbaum bounds test how sensitive your results are to unobserved confounders. DAGs are now the language of causal inference — use them to argue which variables to control for and which to exclude.",
     "Rosenbaum & Rubin (1983) original PSM paper | MatchIt package (R) | https://cran.r-project.org/web/packages/MatchIt/ | Causal Inference: The Mixtape | https://mixtape.scunning.com/05-matching_and_subclassification",
     "MatchIt documentation | https://kosukeimai.github.io/MatchIt/ | cobalt for balance plots | https://cran.r-project.org/web/packages/cobalt/ | DAGitty for drawing DAGs | https://www.dagitty.net/",
     "Dataset: IHDS — impact of microfinance access on household income. Compare households with and without microfinance loan, controlling for confounders via PSM.",
     "Using IHDS: (1) Draw a DAG (on paper or DAGitty) for the causal path from microfinance access to income, including confounders; (2) Estimate propensity score via logistic regression; (3) Nearest-neighbour match 1:1; (4) Check covariate balance using SMD before/after matching (cobalt); (5) Estimate ATT on log income; (6) Run Rosenbaum sensitivity analysis — at what level of hidden bias does your result become insignificant?",
     "Read: Heckman, J., Ichimura, H. & Todd, P. (1998). 'Matching as an Econometric Evaluation Estimator.' Review of Economic Studies. — the foundational paper for modern matching.",
     "potential outcomes, ATE/ATT, DAG, propensity score, covariate balance, SMD, Rosenbaum bounds"),

    (7, 4, "M4: Spatial Econometrics", "Spatial Dependence and Spatial Lag Model",
     "Tobler's First Law; spatial weights matrix (queen/rook contiguity, k-nearest neighbours, distance decay); Moran's I (global and local); spatial lag model (SLM/SAR); interpretation of spatial lag coefficient (ρ); maximum likelihood estimation; spdep and spatialreg packages",
     "Spatial econometrics deals with the fact that OLS assumes observations are independent — but spatial data violates this (neighbouring districts are more similar than distant ones). Ignoring spatial autocorrelation biases standard errors. The spatial lag model adds a spatially lagged version of Y (the outcome averaged across neighbours) as a predictor — the coefficient ρ tells you how strongly a district's outcome is influenced by its neighbours' outcomes. This is directly applicable to your e-waste/circularity research.",
     "Anselin (1988) 'Spatial Econometrics' | LeSage & Pace 'Introduction to Spatial Econometrics' (accessible) | spatialreg package | https://cran.r-project.org/web/packages/spatialreg/",
     "Bivand, Pebesma, Gomez-Rubio 'Applied Spatial Data Analysis with R' | spdep documentation | https://r-spatial.github.io/spdep/ | GeoDa (free spatial analysis GUI) | https://geodacenter.github.io/",
     "Dataset: India district data from SHRUG | https://www.devdatalab.org/shrug_download/ — district-level economic outcomes with spatial structure. Download shapefile from GADM.",
     "Using SHRUG + district shapefile: (1) Create spatial weights matrix (queen contiguity); (2) Compute global Moran's I for nighttime lights — is there spatial autocorrelation?; (3) Make a LISA map identifying clusters; (4) Run OLS of nighttime lights on road density; (5) Run Lagrange Multiplier tests to determine if lag or error model is appropriate; (6) Estimate SLM and interpret ρ; (7) Compare OLS vs SLM coefficients.",
     "Read: Anselin, L. (1995). 'Local Indicators of Spatial Association.' Geographical Analysis, 27(2), 93-115. | https://onlinelibrary.wiley.com/doi/10.1111/j.1538-4632.1995.tb00338.x",
     "spatial weights matrix, Moran's I, spatial lag, ρ, LM tests, spdep, spatialreg"),

    (8, 4, "M4: Spatial Econometrics", "Spatial Error Model and Geographically Weighted Regression",
     "Spatial error model (SEM); SEM vs SLM choice (LM tests); direct, indirect, and total effects in spatial models; Geographically Weighted Regression (GWR): bandwidth selection, local coefficients, mapping coefficient variation; GWmodel package",
     "GWR is a powerful tool for your specific research question: it estimates a separate regression for each location, using a spatial kernel to weight nearby observations more heavily. This reveals where relationships are strongest — e.g., is the effect of formal waste infrastructure on informal sector income spatially heterogeneous? Some districts might show strong displacement effects while others show complementarity. GWR makes these spatial variations visible and quantifiable.",
     "Fotheringham, Brunsdon & Charlton 'Geographically Weighted Regression' (the textbook) | GWmodel package | https://cran.r-project.org/web/packages/GWmodel/",
     "GWR4 software (free) | https://gwr.maynoothuniversity.ie/ | GWmodel vignette | https://cran.r-project.org/web/packages/GWmodel/vignettes/GWmodel.pdf | spatialreg package for SEM",
     "Dataset: Same as Day 7 (SHRUG district data). Estimate GWR and map the local coefficients.",
     "Using SHRUG: (1) Run SEM on your Day 7 specification; compare AIC to SLM; (2) Compute and interpret direct, indirect, and total effects from the SLM; (3) Run GWR with optimal bandwidth; (4) Map local coefficients for road access variable — where is the effect of road access on economic activity strongest?; (5) Map local R² — where does your model fit best/worst? Interpret spatially.",
     "Read: Brunsdon, C., Fotheringham, A.S. & Charlton, M.E. (1996). 'Geographically Weighted Regression.' Geographical Analysis, 28(4), 281-298. | https://onlinelibrary.wiley.com/doi/10.1111/j.1538-4632.1996.tb00936.x",
     "SEM, direct/indirect/total effects, GWR, bandwidth, local coefficients, spatial heterogeneity"),

    (9, 5, "M5: Advanced Methods", "Multilevel / Hierarchical Models",
     "Nested data structures (students in schools, households in districts, districts in states); random intercepts; random slopes; ICC (intraclass correlation); REML estimation; lme4 package in R; visualising random effects; when to use multilevel vs FE",
     "Survey data in India is inherently nested: households are nested within districts, districts within states, states within zones. Ignoring this nesting underestimates standard errors (treats 5000 households in 5 districts as if they were independent). Multilevel models explicitly account for this structure. The ICC (intraclass correlation) tells you what fraction of outcome variance is at the cluster level — if ICC = 0.3, 30% of the variation in child health is between-district rather than between-household.",
     "Gelman & Hill 'Data Analysis Using Regression and Multilevel/Hierarchical Models' (the definitive book) | lme4 package | https://cran.r-project.org/web/packages/lme4/",
     "lme4 vignette | https://cran.r-project.org/web/packages/lme4/vignettes/lmer.pdf | Raudenbush & Bryk 'Hierarchical Linear Models' | Winter (2013): 'A very basic tutorial for performing linear mixed effects analyses in R' | https://arxiv.org/abs/1308.5499",
     "Dataset: NFHS-5 individual data — children nested within mothers nested within districts nested within states. Outcome: child stunting.",
     "Using NFHS child data: (1) Compute ICC for child stunting at district level (using null model); (2) Fit random intercept model: stunting ~ child_char + (1 | district); (3) Fit random slope model: stunting ~ mother_edu + (mother_edu | district) — does the effect of maternal education vary by district?; (4) Plot district-level random intercepts on a map; (5) Compare MLM to OLS with clustered SE: how different are the SE estimates?",
     "Read: Gelman, A. & Hill, J. (2007). 'Data Analysis Using Regression and Multilevel/Hierarchical Models.' Ch.11-12. The most practically useful treatment of multilevel modelling for social scientists.",
     "nested data, ICC, random intercepts, random slopes, REML, lme4, multilevel vs FE"),

    (10, 5, "M5: Advanced Methods", "Quantile Regression",
     "Why OLS misses distributional heterogeneity; quantile regression (Koenker-Bassett); interpretation of quantile regression coefficients; comparing effects across quantiles; rqprocess for testing slope equality across quantiles; quantreg package in R; visualising quantile process",
     "OLS estimates the effect on the mean. Quantile regression estimates the effect at any point of the conditional distribution — the 10th percentile (the poorest), the 50th (median), the 90th (near top). In development economics, you often care specifically about effects on the poorest: does a government program lift the bottom quartile more than the median? Does air pollution harm productivity more at the bottom of the wage distribution? OLS masks these heterogeneous effects.",
     "Koenker & Bassett (1978) founding paper | quantreg package | https://cran.r-project.org/web/packages/quantreg/ | Koenker 'Quantile Regression' (book, 2005)",
     "Roger Koenker's quantreg tutorial | https://cran.r-project.org/web/packages/quantreg/vignettes/rq.pdf | StatQuest: Quantile Regression | Econometrics by Simulation: Quantile Regression | https://www.econometricsbysimulation.com/",
     "Dataset: IHDS household income and education. OR use your earlier econometrics paper data on air quality and labor productivity — does the air pollution effect differ across the wage distribution?",
     "Using IHDS: (1) Run OLS of log income on education, caste, state; (2) Run quantile regression at τ = 0.10, 0.25, 0.50, 0.75, 0.90; (3) Plot quantile process for education coefficient — how does the return to education vary across the income distribution?; (4) Test slope equality: is the slope at τ=0.10 significantly different from τ=0.90?; (5) Interpret: what does this mean for education policy targeting the poorest?",
     "Read: Koenker, R. & Hallock, K. (2001). 'Quantile Regression.' Journal of Economic Perspectives, 15(4), 143-156. | https://www.aeaweb.org/articles?id=10.1257/jep.15.4.143 — very accessible introduction.",
     "quantile regression, τ, conditional distribution, slope heterogeneity, quantreg, quantile process"),

    (11, 6, "M6: Capstone", "Econometrics Capstone: Replication + Extension",
     "Full replication of a published causal identification paper; extension using an additional dataset or method; writing a 500-word methodological critique; identifying threats to validity",
     "This day consolidates everything. Replication forces you to engage with the gap between what a paper claims and what its code actually does. Extension teaches you to ask 'what if' questions: what if we used a different bandwidth in the RDD? What if we included state FE? These are the questions examiners ask at thesis vivas.",
     "Harvard Dataverse | https://dataverse.harvard.edu/ | AEA Data/Code Archive | https://www.aeaweb.org/journals/data",
     "Replication guide | https://social-science-data-editors.github.io/guidance/ | NBER Working Papers for finding replicable papers with India data | https://www.nber.org/",
     "Choose from: (A) Duflo et al. (2003) 'Grandmothers and Granddaughters' — female village councils and girls' education in India (DiD) | (B) Asher & Novosad (2020) 'Rural Roads and Local Economic Development' — PMGSY RDD | https://www.aeaweb.org/articles?id=10.1257/aer.20180268",
     "Replicate the main table of your chosen paper. Then: (1) Try a different bandwidth (if RDD) or control group (if DiD); (2) Add one additional control variable you think is relevant; (3) Test pre-trends or exclusion restriction using an alternative approach; (4) Write 500-word methodological critique discussing: what is the key identification assumption, how convincing is it, and what would be the ideal experiment?",
     "Read: Deaton, A. (2010). 'Instruments, Randomization, and Learning about Development.' Journal of Economic Literature, 48(2), 424-455. | https://www.aeaweb.org/articles?id=10.1257/jel.48.2.424 — a serious critique of the IV and RCT agenda.",
     "Full pipeline: replication, extension, validity threats, methodological critique"),
]

QGIS_DATA = [
    (1, 1, "M1: Spatial Thinking", "What is Spatial Data? Core Concepts",
     "Spatial vs non-spatial data; Tobler's First Law; vector (point, line, polygon) vs raster data; attribute tables; coordinate systems; geographic vs projected CRS; EPSG codes; scale and resolution; metadata",
     "Spatial data has one property that makes it fundamentally different from tabular data: location. Everything else in GIS analysis flows from this. The vector/raster distinction is the first structural choice in any spatial analysis — you choose based on the nature of your phenomenon (discrete objects = vector; continuous surfaces = raster). The CRS is the bridge between a round Earth and flat map — getting it wrong makes your buffers measure in degrees instead of metres, breaking every distance calculation.",
     "A Gentle Introduction to GIS (QGIS Documentation) | https://docs.qgis.org/3.34/en/docs/gentle_gis_introduction/ — read Ch.1-5",
     "National Geographic GIS Explainer | https://education.nationalgeographic.org/resource/geographic-information-system-gis/ | Esri 'What is GIS?' | https://www.esri.com/en-us/what-is-gis/overview | EPSG registry | https://epsg.io/",
     "No software yet. Download Chennai ward shapefile from GADM | https://gadm.org/download_country.html — select India, Level 3 (district). Examine the .shp, .dbf, .prj files in a text editor to understand file structure.",
     "Answer (without software, just thinking): (1) Is 'number of e-waste dealers per ward' vector or raster data? (2) Is 'land surface temperature in summer' vector or raster? (3) If I have a point layer of hospitals and want to know which hospitals are within 2km of a major road — what spatial operation is this? (4) If my shapefile is in WGS84 but I need to compute area in km² — what must I do first?",
     "Read: Longley, P.A. et al. (2015). 'Geographic Information Science and Systems.' Ch.1-2. Focus on conceptual content, skip technical diagrams.",
     "vector/raster, CRS, EPSG, attribute table, projection, Tobler's First Law, metadata"),

    (2, 1, "M1: Spatial Thinking", "QGIS Interface and First Map",
     "QGIS installation and interface overview; Layers Panel; Map Canvas; Toolbars; Browser Panel; Loading vector layers (drag-and-drop and Layer menu); Attribute table operations; Symbology (single symbol, categorised, graduated); Labels; Save as project (.qgz)",
     "QGIS can look overwhelming at first — ignore everything except the five things you need: (1) Layer panel on the left; (2) Map canvas in the centre; (3) Layer menu for adding data; (4) right-click menu on a layer for most operations; (5) the Processing toolbar for analysis. The attribute table is just a spreadsheet — every row is a geographic feature, every column is a variable you can use for styling or analysis. Always save your project (.qgz) — it saves your layer configuration but not the data files.",
     "QGIS Training Manual Module 1-3 | https://docs.qgis.org/3.34/en/docs/training_manual/",
     "QGIS.org tutorial videos | https://www.qgis.org/resources/hub/ | Klas Karlsson QGIS beginners | https://youtube.com/playlist?list=PLuSaaR-OElIq1Sq85RV5aFwq4FPDUfLj9 | ESRI Mapping resources | https://learn.arcgis.com/",
     "Dataset: India district boundaries from GADM (download Level 2) | https://gadm.org/download_country.html | Census 2011 district data from censusindia.gov.in or pre-cleaned from datameet.org | https://github.com/datameet",
     "Load India district shapefile. Style the map: (1) Use graduated colours to show a census variable (e.g., literacy rate, female-male ratio, population density); (2) Add district name labels for districts in Tamil Nadu only (use rule-based labelling, filter by STATE = 'Tamil Nadu'); (3) Add a basemap using QuickMapServices plugin; (4) Export a print-ready PNG with title, legend, north arrow, and scale bar.",
     "None",
     "QGIS interface, layers panel, symbology, graduated/categorised, labels, QuickMapServices, print layout"),

    (3, 2, "M2: Spatial Operations", "Geoprocessing: Buffer, Clip, Dissolve, Intersect",
     "Buffer (fixed, variable); Clip (cookie-cutter operation); Dissolve (merge polygons by attribute); Union; Difference; Intersect; Spatial selection (Select by Location); running tools from Processing Toolbox; saving outputs as geopackage",
     "These are the five operations that handle 80% of all spatial analysis tasks. Buffer: 'what is within X distance of a hospital?' Clip: 'give me only the roads within Chennai's boundary.' Dissolve: 'merge all ward polygons belonging to the same district into one district polygon.' Intersect: 'find areas where industrial zones overlap with low-income residential areas.' Always save intermediate outputs as GeoPackage (.gpkg) — it is a modern, single-file format better than the old Shapefile format (which splits data into 4+ files).",
     "QGIS Training Manual Module 4 | https://docs.qgis.org/3.34/en/docs/training_manual/vector_analysis/index.html | QGIS Processing Toolbox documentation",
     "Spatial Analysis tutorial (GIS Geography) | https://gisgeography.com/spatial-analysis/ | ESRI Buffer explainer | https://pro.arcgis.com/en/pro-app/latest/tool-reference/analysis/buffer.htm",
     "Dataset: OpenStreetMap Chennai data | Download from Geofabrik | https://download.geofabrik.de/asia/india.html — contains roads, buildings, amenities as separate layers. Also: Chennai ward boundaries.",
     "Using Chennai OSM data: (1) Buffer all hospitals by 2km — create a 'hospital catchment' layer; (2) Clip the road network to only roads within Chennai city boundary; (3) Dissolve ward boundaries to district level using ward attribute; (4) Intersect residential areas with the 2km hospital buffers — what proportion of residential area is within 2km of a hospital?; (5) Save all outputs as .gpkg layers.",
     "None",
     "buffer, clip, dissolve, union, intersect, spatial selection, GeoPackage"),

    (4, 2, "M2: Spatial Operations", "Spatial Joins and Table Operations",
     "Join by location (spatial join); join by attribute (table join); adding CSV data to a shapefile; field calculator; expression builder; creating centroids; counting points in polygons; statistics by polygon",
     "Spatial join is the operation that links your fieldwork data to geographic context. You geocode your interview locations → you have points. You load ward polygons. Spatial join adds ward attributes to each point. Now your interview data knows which ward it is in, what the ward's population density is, how far it is from formal waste infrastructure. The field calculator extends this: compute new variables directly in QGIS using expressions.",
     "QGIS Spatial Join documentation | https://docs.qgis.org/3.34/en/docs/user_manual/processing_algs/qgis/vectorgeneral.html#join-attributes-by-location",
     "QGIS training: Table Joins | https://docs.qgis.org/3.34/en/docs/training_manual/vector_analysis/table_joins.html | GIS Geography spatial join | https://gisgeography.com/spatial-join/",
     "Dataset: Create a CSV of 20 geocoded locations in Chennai (use Google Maps to get coordinates of real places — markets, schools, industrial areas) + Census ward data.",
     "Using your geocoded CSV: (1) Import CSV as a point layer (Layer → Add Delimited Text Layer); (2) Spatial join: add ward attributes to each point; (3) Join by attribute: merge literacy rate from a CSV table to your ward polygon using district code as the common key; (4) Field calculator: compute population density (pop/area) for each ward; (5) Count points in polygons: how many of your 20 locations fall in each ward?",
     "None",
     "spatial join, table join, field calculator, centroid, count points in polygon, expression builder"),

    (5, 3, "M3: Raster Analysis", "Introduction to Raster Data",
     "Raster data model; bands; resolution (spatial, temporal, spectral); loading raster in QGIS; raster symbology (singleband, multiband, hillshade); raster calculator; zonal statistics; clipping raster to study area; resampling",
     "Rasters represent continuous phenomena — elevation, temperature, rainfall, vegetation density, nighttime lights — as grids of values. For your research: satellite-derived nighttime lights (a proxy for economic activity) and land surface temperature (for urban heat island analysis) are rasters. The raster calculator lets you do algebra on rasters: (Band1 - Band2) / (Band1 + Band2) = NDVI (vegetation index). Zonal statistics computes raster statistics within each polygon — e.g., mean LST within each ward.",
     "QGIS Raster Analysis documentation | https://docs.qgis.org/3.34/en/docs/training_manual/rasters/index.html | Earth Lab raster intro | https://www.earthdatascience.org/courses/intro-to-earth-data-science/file-formats/use-spatial-data/use-raster-data/",
     "NASA Earthdata (free satellite data) | https://www.earthdata.nasa.gov/ | Copernicus Open Access Hub | https://browser.dataspace.copernicus.eu/ | USGS Earth Explorer | https://earthexplorer.usgs.gov/",
     "Dataset: VIIRS Nighttime Lights (Black Marble) for India | https://ladsweb.modaps.eosdis.nasa.gov/missions-and-measurements/products/VNP46A3/ — free, monthly composite, 500m resolution. Register on NASA Earthdata.",
     "Using VIIRS nighttime lights raster for India: (1) Load and style the raster; (2) Clip to Tamil Nadu using 'Clip raster by mask layer'; (3) Run 'Zonal Statistics' to compute mean nighttime radiance per district; (4) Join results back to district shapefile; (5) Make a choropleth map of nighttime lights by district as a proxy for economic activity; (6) Identify top-5 and bottom-5 districts by nighttime lights.",
     "None",
     "raster model, bands, resolution, raster calculator, zonal statistics, clip raster, NDVI, VIIRS"),

    (6, 3, "M3: Raster Analysis", "Remote Sensing Basics — Satellite Imagery",
     "Electromagnetic spectrum; multispectral vs hyperspectral imagery; Landsat and Sentinel-2 sensors; band combinations (true colour, false colour, SWIR); NDVI (vegetation); NDBI (built-up index); NDWI (water); land cover classification (supervised basics); downloading from Copernicus/USGS; Google Earth Engine (intro)",
     "Sentinel-2 (European Space Agency) provides 10m resolution, 13-band multispectral imagery, freely available every 5 days. For urban research: NDVI shows vegetation; NDBI (Normalized Difference Built-up Index) shows extent of built-up area; land surface temperature from Landsat 8/9 Band 10 shows urban heat islands. These are directly useful for characterising the physical environment of urban informal areas.",
     "Copernicus Open Access Hub | https://browser.dataspace.copernicus.eu/ | USGS Earth Explorer | https://earthexplorer.usgs.gov/ | Google Earth Engine (requires registration) | https://earthengine.google.com/",
     "Sentinel Hub EO Browser | https://www.sentinel-hub.com/explore/eobrowser/ | NASA Applied Remote Sensing Training | https://appliedsciences.nasa.gov/what-we-do/capacity-building/arset | ESA Sentinel-2 user guide | https://sentinel.esa.int/web/sentinel/user-guides/sentinel-2-msi",
     "Dataset: Sentinel-2 image over Chennai | Download from Copernicus Browser — search for a recent cloud-free image over Chennai and download Bands 4,3,2 (true colour) + Band 8 (NIR for NDVI).",
     "Using Sentinel-2 over Chennai: (1) Display true colour composite (RGB = B4,B3,B2); (2) Display false colour composite (NIR-R-G = B8,B4,B3) — vegetation appears bright red; (3) Compute NDVI using raster calculator: (B8-B4)/(B8+B4); (4) Compute NDBI: (SWIR-NIR)/(SWIR+NIR) using bands B11 and B8; (5) Classify NDVI into: water (<0), barren (0-0.2), sparse vegetation (0.2-0.4), dense vegetation (>0.4); (6) Compute area of each class within Chennai boundary.",
     "Read: Jensen, J.R. (2016). 'Introductory Digital Image Processing.' Ch.1-2 — foundational remote sensing concepts.",
     "multispectral, Landsat, Sentinel-2, NDVI, NDBI, land cover, band combinations, remote sensing"),

    (7, 4, "M4: Cartography", "Map Design and Geovisualisation Principles",
     "Principles of cartographic design: figure-ground, hierarchy, contrast, balance; colour theory for maps: sequential, diverging, qualitative palettes; colourblind-safe palettes; choosing classification schemes (equal interval, quantile, natural breaks/Jenks, standard deviation); QGIS Print Layout: map frames, legends, scalebars, north arrows, inset maps; exporting at 300 dpi",
     "A bad map misleads even with accurate data. Classification scheme choice is a hidden design decision: quantile breaks always show variation even when variation is small; equal intervals can make everything look similar when a few extreme values dominate. For social science maps: always use colourblind-safe palettes (ColorBrewer), always state your classification scheme in the legend, always include sample size context. The map is an argument — every design choice is a rhetorical choice.",
     "ColorBrewer (essential tool) | https://colorbrewer2.org/ | Cartography Guide (Axis Maps) | https://www.axismaps.com/guide | QGIS Print Layout tutorial | https://docs.qgis.org/3.34/en/docs/training_manual/map_composer/index.html",
     "Mapschool.io (conceptual) | https://mapschool.io/ | Fundamentals of Data Visualization (Wilke, free) Ch.15 | https://clauswilke.com/dataviz/ | ESRI Cartography resources | https://www.esri.com/arcgis-blog/products/arcgis-pro/mapping/",
     "Dataset: Use your NFHS district data + India shapefile. Create a multi-panel policy map.",
     "Create a publication-quality 4-panel map figure: (A) Under-5 mortality by district (sequential palette, natural breaks); (B) Change in under-5 mortality between NFHS-4 and NFHS-5 (diverging palette, 0 = midpoint); (C) Toilet access coverage (sequential, quantile breaks); (D) Inset map of Tamil Nadu zoomed in. Add: title, subtitle, legend, source, north arrow, scale bar. Export at 300 dpi. Write 100-word caption.",
     "None",
     "figure-ground, colour theory, ColorBrewer, classification schemes, Jenks, print layout, 300dpi"),

    (8, 4, "M4: Cartography", "QGIS Capstone: Chennai E-Waste Spatial Analysis",
     "Full spatial analysis workflow: data acquisition, cleaning, geocoding, joining, analysis, cartography; applying all operations from previous days to a coherent research question",
     "This is the day where all previous QGIS skills integrate into a research-quality output. The goal is to produce a map figure you could put directly into a thesis chapter — with methods described precisely enough to be replicated.",
     "All previous resources.",
     "All previous resources. SHRUG documentation | https://www.devdatalab.org/shrug",
     "Data to collect: (1) Chennai ward/division shapefile; (2) NFHS-5 ward-level sanitation and poverty indicators where available, else district level; (3) Geocode 15+ known e-waste / electronics market locations (Ritchie Street, SP Road, etc.) using Google Maps; (4) OpenStreetMap roads layer for Chennai.",
     "Produce a thesis-ready map: (1) Choropleth of district-level development indicator (poverty or sanitation from NFHS); (2) Point layer of e-waste/electronics market locations; (3) 2km buffer around each market; (4) Road network for context; (5) Style in QGIS, export at 300dpi with full cartographic elements. Write 200-word 'Spatial Data and Methods' section describing exactly what you did.",
     "Read: Ghosh, S. et al. (2020). 'Mapping India's e-waste.' Resources, Conservation and Recycling, 161. — examine how they describe their spatial methods and what maps they include.",
     "Full workflow, research-quality map, methods documentation, cartographic standards"),
]

SPATIAL_PY_DATA = [
    (1, 1, "M1: GeoPandas Foundations", "GeoDataFrame Basics",
     "GeoDataFrame structure; reading shapefiles and GeoJSON; CRS management (to_crs(), set_crs()); geometry types; basic plotting (plot()); exploring geometry columns; converting between CRS; creating GeoDataFrame from lat/lon columns",
     "GeoPandas extends Pandas by adding a geometry column — everything you already know about Pandas operations (filter, groupby, merge) works exactly the same way on a GeoDataFrame. The geometry column holds the spatial information (points/lines/polygons) as Shapely geometry objects. The critical skill is CRS management: always check .crs after loading, always reproject to a metric CRS before any distance/area calculations.",
     "GeoPandas documentation | https://geopandas.org/en/stable/ | GeoPandas Introduction tutorial | https://geopandas.org/en/stable/getting_started/introduction.html",
     "Python for Spatial Analysis (book) | Spatial Data Science (free online) | https://geographicdata.science/book/ | GeoPandas user guide | https://geopandas.org/en/stable/docs/user_guide.html",
     "Dataset: India district boundaries from GADM | Download GeoJSON version (easier than shapefile for GeoPandas): go to https://gadm.org/download_country.html — India, GeoJSON, Level 2",
     "Using India GeoJSON: (1) Load with gpd.read_file(); (2) Print .crs — what CRS is it in?; (3) Filter to only Tamil Nadu districts; (4) Reproject to UTM Zone 44N (EPSG:32644); (5) Compute area of each district in km² and add as a column; (6) Plot a choropleth of area; (7) Create a GeoDataFrame from a list of 5 lat/lon pairs (use locations you know in Chennai) and plot them on top of the district map.",
     "None",
     "GeoDataFrame, geometry column, CRS, to_crs, Shapely, read_file, choropleth"),

    (2, 1, "M1: GeoPandas Foundations", "Spatial Operations in GeoPandas",
     "buffer(); dissolve(); clip(); overlay() (intersection, union, difference, symmetric difference); sjoin() spatial join; distance(); centroid; convex hull; unary_union; bounding box; within(), contains(), intersects()",
     "GeoPandas spatial operations are vectorised — they apply to all features in a GeoDataFrame at once. sjoin() is the spatial equivalent of pd.merge() — it joins two GeoDataFrames based on a spatial relationship (within, intersects, contains). Always verify your CRS is projected (not geographic) before buffer() and distance() — otherwise distances are in degrees, which is meaningless.",
     "GeoPandas Geometric Manipulations | https://geopandas.org/en/stable/docs/user_guide/geometric_manipulations.html | GeoPandas Set Operations | https://geopandas.org/en/stable/docs/user_guide/set_operations.html",
     "Spatial Data Science book Ch.3-4 | https://geographicdata.science/book/ | PySAL documentation | https://pysal.org/",
     "Dataset: Chennai OSM data (roads, hospitals, markets) downloaded from Geofabrik | https://download.geofabrik.de/asia/india-latest.osm.pbf — filter to Tamil Nadu using osmium or download Chennai extract.",
     "Using Chennai OSM data: (1) Buffer all hospitals by 2km using buffer(); (2) Dissolve hospitals layer by hospital type; (3) Clip roads to Chennai boundary; (4) Spatial join: assign each hospital to its district using sjoin(); (5) Count hospitals per district using groupby after sjoin; (6) Find all markets within 500m of a major road using within(); (7) Compute distance matrix between 5 selected locations.",
     "None",
     "buffer, dissolve, clip, overlay, sjoin, within, distance, centroid, unary_union"),

    (3, 2, "M2: Geocoding + Data Enrichment", "Geocoding and Data from APIs in Python",
     "geopy geocoding (Nominatim, GoogleV3); batch geocoding with rate limiting; reverse geocoding; OSMnx for street networks; downloading OSM data programmatically; Overpass API; populating research datasets with spatial coordinates",
     "Geocoding is how you convert 'Ritchie Street, Chennai' into (13.0606, 80.2582). For a fieldwork-based thesis, this is how your qualitative interview locations become spatial data. The Nominatim geocoder (based on OpenStreetMap) is free but rate-limited to 1 request/second — for large batches, add time.sleep(1) between requests. OSMnx downloads complete street networks from OpenStreetMap, letting you do network analysis (shortest path, accessibility) programmatically.",
     "geopy documentation | https://geopy.readthedocs.io/ | OSMnx documentation | https://osmnx.readthedocs.io/ | Overpass API | https://overpass-turbo.eu/",
     "OSMnx examples | https://github.com/gboeing/osmnx-examples | Geopy tutorial | https://towardsdatascience.com/geocoding-in-python-using-geopy-2a7bf28b7f52 | Overpass Turbo (visual query builder) | https://overpass-turbo.eu/",
     "Dataset: Your own fieldwork locations OR a list of 20 well-known locations in Chennai (hospitals, markets, informal settlements, government offices — use public knowledge).",
     "Build a spatial dataset from scratch: (1) Create a list of 20 location names in Chennai; (2) Geocode all of them using Nominatim with proper rate limiting; (3) Create a GeoDataFrame; (4) Use OSMnx to download the Chennai street network; (5) Plot your locations on the street network; (6) Use OSMnx to find the 3 nearest hospitals to each of your 20 locations.",
     "None",
     "geocoding, Nominatim, rate limiting, reverse geocoding, OSMnx, Overpass API, street networks"),

    (4, 2, "M2: Geocoding + Data Enrichment", "Raster Analysis with Rasterio",
     "rasterio.open(); reading bands; CRS and transform; masking and clipping raster to vector; raster arithmetic (NDVI calculation); zonal statistics with rasterstats; converting raster to array for numpy operations; saving raster outputs",
     "Rasterio is Python's primary raster library — it treats rasters as NumPy arrays, so all your NumPy skills apply. The key concept: each pixel has both a value and a location (defined by the transform matrix). zonal_stats from rasterstats computes statistics of a raster within each polygon — this is how you extract mean nighttime light intensity, mean NDVI, or mean land surface temperature for each district/ward.",
     "Rasterio documentation | https://rasterio.readthedocs.io/ | rasterstats documentation | https://pythonhosted.org/rasterstats/",
     "Earth Lab Python raster tutorial | https://www.earthdatascience.org/courses/use-data-open-source-python/intro-raster-data-python/ | Automating GIS Processes (free course) | https://autogis-site.readthedocs.io/",
     "Dataset: VIIRS Nighttime Lights (Monthly Composite) for India | https://ladsweb.modaps.eosdis.nasa.gov/ — or download a pre-processed version from: https://www.kaggle.com/datasets/tunguz/nighttime-lights — Sentinel-2 tile over Chennai from Copernicus Browser.",
     "Using VIIRS nighttime lights: (1) Open with rasterio, print metadata (CRS, resolution, bounds); (2) Clip to India extent; (3) Compute zonal statistics (mean radiance) for each state using rasterstats; (4) Join to state GeoDataFrame; (5) Plot choropleth of mean nighttime radiance by state; (6) Using Sentinel-2 data: compute NDVI using rasterio array operations; (7) Run zonal statistics to get mean NDVI per ward.",
     "None",
     "rasterio, bands, transform, masking, NDVI, zonal statistics, rasterstats, NumPy arrays"),

    (5, 3, "M3: Interactive + Advanced", "Interactive Maps and Web Visualisation",
     "Folium (Leaflet.js wrapper); markers, popups, tooltips; choropleth maps; heatmaps; plugin layers; Plotly Express choropleth; Kepler.gl for exploratory analysis; saving as HTML",
     "Interactive maps let you zoom, click on features to see data, and toggle layers — invaluable for exploratory analysis of spatial data. For a thesis or policy presentation: embed a Folium map in a Jupyter notebook to explore your data interactively before finalising your static QGIS maps. Kepler.gl handles large datasets (100,000+ points) with animated time sliders, useful for showing how informal markets change over time if you have longitudinal data.",
     "Folium documentation | https://python-visualization.github.io/folium/ | Plotly geographic plots | https://plotly.com/python/maps/ | Kepler.gl | https://kepler.gl/",
     "Towards Data Science: Folium tutorial | https://towardsdatascience.com/creating-choropleth-maps-with-pythons-folium-library-cfacfb40f56a | Plotly Maps documentation | https://plotly.com/python/choropleth-maps/",
     "Dataset: Your Chennai data from previous days (points + ward polygons + development indicators).",
     "Create: (1) A Folium map of Chennai with: ward choropleth (NFHS indicator), point markers for your geocoded locations with popup showing location name and attributes, hospital catchment buffers as a toggleable layer; (2) A Plotly choropleth of India states by a WDI indicator with hover tooltips; (3) Save as standalone HTML files. Bonus: animate nighttime lights over time using Folium TimestampedGeoJson.",
     "None",
     "Folium, choropleth, popups, layer control, Plotly, Kepler.gl, interactive HTML"),

    (6, 3, "M3: Interactive + Advanced", "Spatial Python Capstone",
     "End-to-end spatial analysis pipeline in Python: acquisition, cleaning, geocoding, spatial operations, raster analysis, visualisation, and reporting in a Jupyter notebook",
     "This day integrates all Spatial Python skills into a coherent research workflow in a single Jupyter notebook. The deliverable is a notebook that someone else could clone and run end-to-end with no additional explanation — the gold standard for reproducible spatial research.",
     "All previous resources.",
     "Spatial Data Science book | https://geographicdata.science/book/ — read the case study chapters.",
     "All data from previous days.",
     "Build a complete spatial analysis notebook titled 'Spatial Profile of E-Waste and Urban Development in Chennai': (1) Load ward boundaries and census data; (2) Geocode your fieldwork locations; (3) Compute 500m and 1km service area buffers; (4) Spatial join wards to field locations; (5) Compute zonal stats of NDVI (green cover) and nighttime lights per ward; (6) Run a simple OLS regression of a development indicator on distance to nearest formal waste facility and NDVI; (7) Produce 3 maps: context, field sites, spatial regression setup. Export as HTML.",
     "Read: Rey, S., Arribas-Bel, D. & Wolf, L. (2023). 'Geographic Data Science with Python.' Preface and Ch.1. | https://geographicdata.science/book/ — the definitive Python spatial data science textbook.",
     "Full pipeline, reproducibility, Jupyter, spatial regression, maps, HTML export"),
]

SPATIAL_R_DATA = [
    (1, 1, "M1: sf Package", "Spatial Data in R with sf",
     "sf package philosophy; st_read(), st_write(); sf object structure (data frame + geometry list column); st_crs(), st_transform(); st_geometry(), st_bbox(); basic operations: st_area(), st_length(), st_distance(), st_centroid(); plot.sf(); simple_features for geometry types",
     "sf (Simple Features) is the modern standard for vector spatial data in R. Its key design: a spatial object is just a data frame with an extra 'geometry' column — so all tidyverse operations (filter, mutate, group_by, etc.) work on it directly. The 'sf' abbreviation stands for 'Simple Features', which is an ISO standard for representing geographic features — making sf interoperable with databases, Python (via PyGEOS), and GIS software.",
     "Geocomputation with R Ch.1-2 (free) | https://geocompr.robinlovelace.net/ | sf package documentation | https://r-spatial.github.io/sf/",
     "sf cheatsheet | https://github.com/rstudio/cheatsheets/blob/main/sf.pdf | r-spatial blog | https://r-spatial.org/ | Edzer Pebesma's sf tutorials | https://cran.r-project.org/web/packages/sf/vignettes/sf1.html",
     "Dataset: India districts from GADM (download .rds file for R directly) | https://gadm.org/download_country.html | OR: use rnaturalearth package: ne_states(country='India', returnclass='sf')",
     "Using India spatial data in sf: (1) Load India states using rnaturalearth; (2) Check and transform CRS to EPSG:32644 (UTM); (3) Compute area of each state in km²; (4) Filter to South India (specific states using filter()); (5) Compute centroid for each state; (6) Compute distance matrix between state centroids; (7) Plot: fill states by area, add centroid points, label top-5 largest states.",
     "None",
     "sf, st_read, st_transform, st_area, st_centroid, st_distance, geometry list column, simple features"),

    (2, 1, "M1: sf Package", "Spatial Operations in R",
     "st_buffer(); st_union(); st_intersection(); st_difference(); st_join() (left, inner, nearest); st_within(); st_contains(); st_intersects(); st_nn() for nearest neighbour; st_snap(); nngeo package for k-nearest",
     "st_join() is the R equivalent of a spatial join — it transfers attributes from one layer to another based on spatial relationship. The predicate argument specifies the relationship: st_within (point falls inside polygon), st_intersects (geometries share any space), st_nearest_feature (useful for finding nearest facility). Always use left = TRUE to keep all features from the left table even when no match is found (preserving your analytical universe).",
     "Geocomputation with R Ch.4-5 | https://geocompr.robinlovelace.net/spatial-operations.html | sf spatial operations vignette | https://r-spatial.github.io/sf/articles/sf3.html",
     "Geometric operations vignette | https://r-spatial.github.io/sf/articles/sf5.html | nngeo package | https://cran.r-project.org/web/packages/nngeo/",
     "Dataset: Chennai OSM data loaded into R via osmdata package. install.packages('osmdata'). hospitals <- opq('Chennai, India') |> add_osm_feature(key = 'amenity', value = 'hospital') |> osmdata_sf()",
     "Using osmdata: (1) Download hospitals and markets in Chennai; (2) Buffer hospitals by 2km; (3) Spatial join: which ward does each hospital fall in?; (4) Count hospitals per ward using group_by; (5) Find the 3 nearest hospitals for each market using nngeo::st_nn(); (6) Compute proportion of total ward area within 2km of a hospital; (7) Map: ward choropleth of hospital accessibility index.",
     "None",
     "st_buffer, st_union, st_join, st_within, nngeo, osmdata, spatial predicates"),

    (3, 2, "M2: tmap Visualisation", "Static and Interactive Maps with tmap",
     "tmap package philosophy; tm_shape() + tm_polygons/tm_dots/tm_lines; aesthetic mappings (col, size, alpha, border.col); tm_layout(), tm_legend(), tm_compass(), tm_scale_bar(); tmap_mode('view') for interactive; faceted maps with tm_facets(); tmap_arrange() for multi-panel; saving with tmap_save()",
     "tmap's design mirrors ggplot2 — it uses a layered approach where you specify the shape (spatial data) and then the visual layer (polygon fill, dot size, etc.). The mode switch between static ('plot') and interactive ('view') is seamless — the same code produces either output. For thesis work: develop your map in interactive mode for exploration, then switch to plot mode for the final print version.",
     "tmap documentation | https://r-tmap.github.io/tmap/ | tmap vignette | https://cran.r-project.org/web/packages/tmap/vignettes/tmap-getstarted.html",
     "Geocomputation with R Ch.8 (Mapping) | https://geocompr.robinlovelace.net/adv-map.html | tmap book (in development) | https://r-tmap.github.io/tmap-book/",
     "Dataset: India/Tamil Nadu district data + NFHS-5 district indicators.",
     "Create using tmap: (1) A static choropleth of Tamil Nadu districts by under-5 mortality (natural breaks, Blues palette, with compass, scalebar); (2) A bivariate map of two variables using biscale package (challenging but powerful); (3) A faceted map showing the same indicator across 4 NFHS rounds where available; (4) Switch to interactive mode and add pop-up labels; (5) Save at 300dpi as PNG and as interactive HTML.",
     "None",
     "tmap, tm_shape, tm_polygons, tm_dots, tmap_mode, tm_facets, tmap_arrange, biscale"),

    (4, 2, "M2: tmap Visualisation", "Spatial Statistics in R: Moran's I and LISA",
     "spdep package: poly2nb(), nb2listw(), listw2mat(); global Moran's I (moran.test()); Moran's I permutation test (moran.mc()); Getis-Ord Gi* (hotspot analysis); Local Moran's I (localmoran()); LISA cluster classification and mapping; spatial correlograms",
     "Spatial autocorrelation analysis is the bridge between descriptive spatial mapping and inferential spatial analysis. The question is: is the spatial pattern you see on your map statistically non-random? Global Moran's I answers this globally. LISA maps answer it locally — identifying the specific locations where clustering or dispersion occurs. Gi* (Getis-Ord) identifies hot spots and cold spots, accounting for spatial scale. For your thesis: a LISA map of informal e-waste activity showing spatial clustering is a direct, quantifiable spatial finding.",
     "spdep documentation | https://r-spatial.github.io/spdep/ | Anselin (1995) LISA paper | Geocomputation with R Ch.12 | https://geocompr.robinlovelace.net/",
     "spdep vignette | https://cran.r-project.org/web/packages/spdep/vignettes/spdep.pdf | OpenGeoda spatial analysis resources | https://geodacenter.github.io/workbook/6a_local_auto/lab6a.html | GeoDa (GUI version for validation) | https://geodacenter.github.io/",
     "Dataset: SHRUG district data + India district shapefile. Analyse spatial clustering of nighttime lights (economic activity proxy).",
     "Using SHRUG district data: (1) Create queen contiguity weights; (2) Moran scatter plot (moran.plot()); (3) Permutation-based Moran's I test with 999 simulations; (4) Local Moran's I for each district; (5) Classify each district as HH/HL/LH/LL/NS (not significant); (6) Map LISA clusters — where are the high-high clusters? Interpret substantively; (7) Gi* hotspot analysis; (8) Correlogram showing how Moran's I decays with spatial lag distance.",
     "Read: Anselin, L. (1995). 'Local Indicators of Spatial Association.' Geographical Analysis, 27(2), 93-115. | https://onlinelibrary.wiley.com/doi/10.1111/j.1538-4632.1995.tb00338.x",
     "spdep, poly2nb, nb2listw, moran.test, moran.mc, localmoran, LISA, Gi*, correlogram"),

    (5, 3, "M3: Spatial Regression", "Spatial Regression in R",
     "spatialreg package: lagsarlm() for spatial lag; errorsarlm() for spatial error; LM tests for model selection (lm.LMtests()); spatial Durbin model; impacts(): direct, indirect, total effects; spatial two-stage least squares; model comparison (AIC, log-likelihood)",
     "Spatial regression accounts for the fact that OLS residuals are spatially autocorrelated — violating the independence assumption. The choice between spatial lag model (spatial spillover in the outcome — neighbours' outcomes affect your outcome) and spatial error model (spatial clustering in unmeasured factors — neighbours share unobserved determinants) is guided by the LM tests. The interpretation of the spatial lag model requires computing direct and indirect (spillover) effects — a change in X in one district affects not just that district but also its neighbours through the spatial lag.",
     "spatialreg package | https://cran.r-project.org/web/packages/spatialreg/ | LeSage & Pace 'Introduction to Spatial Econometrics' Ch.1-3 | Spatial Econometrics with R (Bivand) | https://spatialreg.r-forge.r-project.org/",
     "Geocomputation with R: Spatial modelling | https://geocompr.robinlovelace.net/ | R-spatial ecosystem overview | https://r-spatial.org/r/2018/04/09/ie3.html | spdep vignette",
     "Dataset: SHRUG data + district shapefile. Outcome: nighttime lights. Predictor: road access, literacy, urban share.",
     "Full spatial regression exercise: (1) Run OLS; (2) Test residuals for spatial autocorrelation (Moran's I on residuals); (3) Run LM tests (lm.LMtests()) — which model does the data suggest?; (4) Run spatial lag model; (5) Run spatial error model; (6) Compare AIC; (7) Compute impacts() from spatial lag model: report direct, indirect, and total effect of road access; (8) Interpret: by how much does improving road access in one district affect nighttime lights in neighbouring districts?",
     "Read: LeSage, J. & Pace, R.K. (2009). 'Introduction to Spatial Econometrics.' Ch.1-2. For a short intro: Elhorst, J.P. (2010). 'Applied Spatial Econometrics: Raising the Bar.' Spatial Economic Analysis, 5(1).",
     "spatialreg, lagsarlm, errorsarlm, LM tests, direct/indirect/total effects, spatial Durbin, AIC"),

    (6, 3, "M3: Spatial Regression", "Spatial R Capstone: Replicating an IOER-Style Analysis",
     "Full spatial analysis pipeline in R: data acquisition, spatial join, Moran's I, LISA mapping, spatial regression, interpretation, and report generation in R Markdown",
     "This final day produces a complete, publication-ready spatial analysis in R Markdown — the kind of analysis you would present to IOER researchers. The goal is to demonstrate that you can run an analysis from raw data to interpretable, map-illustrated findings in a reproducible format.",
     "All previous resources. Geocomputation with R full text | https://geocompr.robinlovelace.net/",
     "IOER Monitor methodology documentation | https://www.ioer-monitor.de/en/methodology/ — understand how IOER constructs its spatial indicators.",
     "SHRUG (economic panel), GADM shapefile, NFHS district data, osmdata (amenities). All previously used.",
     "Create an R Markdown report: 'Spatial Analysis of Urban Development Patterns in Tamil Nadu'. Include: (1) Data section: load and join SHRUG + NFHS + OSM data; (2) Descriptive maps: 3 choropleth maps; (3) Spatial autocorrelation: Moran's I and LISA map; (4) Regression: OLS, then appropriate spatial model; (5) Impacts table; (6) Discussion: 300-word interpretation of spatial spillover findings; (7) Knit to PDF. This document is the template for a thesis chapter.",
     "Read: Bivand, R. (2022). 'R Packages for Analyzing Spatial Data: A Comparative Case Study with Areal Data.' Geographical Analysis, 54(3). — a review of the full spatial R ecosystem from one of its main architects.",
     "Full pipeline, R Markdown, spatial regression, impacts, LISA, reproducible report"),
]

# ══════════════════════════════════════════════════════════════════════════════
#  OVERVIEW DATA
# ══════════════════════════════════════════════════════════════════════════════
OVERVIEW_DATA = [
    ("Module 0", "CS & Programming Foundations", 8, 16, "None — start here",
     "How computers work, the command line, algorithms, data structures, git version control, SQL/databases, and how all research software tools fit together",
     "Confident terminal user, git-tracked thesis project, SQL basics, understanding of why code behaves the way it does"),
    ("Module 1", "Python Foundations", 14, 28, "Module 0 (helpful, not required)",
     "Complete Python programming from scratch for research: syntax, data structures, Pandas, NumPy, visualisation, APIs, text analysis, reproducibility",
     "Jupyter Notebook, competent data manipulation, reproducible research workflow"),
    ("Module 2", "R Language", 10, 20, "Module 0 + basic programming literacy",
     "R and Tidyverse for data analysis: dplyr, tidyr, ggplot2, R Markdown, statistical testing, regression, functional programming",
     "R Markdown reports, ggplot2 publication charts, dplyr pipelines"),
    ("Module 3", "Data Analysis & Statistical Modelling", 10, 20, "Python or R basics",
     "Statistical foundations: distributions, hypothesis testing, regression, logistic/count models, time series, index construction, Bayesian basics",
     "Regression analysis, index construction, appropriate model selection"),
    ("Module 4", "Advanced Econometrics", 11, 22, "Stats Modelling module",
     "Causal inference toolkit: OVB, IV/2SLS, panel FE/RE, DiD, RDD, PSM/matching, spatial econometrics, GWR, multilevel models, quantile regression",
     "Causal identification, spatial econometrics, replication exercise"),
    ("Module 5", "QGIS & Spatial Thinking", 8, 16, "None (conceptual foundations first)",
     "GIS from zero: spatial data concepts, QGIS operations, geoprocessing, raster analysis, remote sensing, cartography",
     "Publication-quality maps, raster/vector analysis, Chennai spatial dataset"),
    ("Module 6", "Spatial Python (GeoPandas)", 6, 12, "Module 1 + Module 5",
     "Spatial analysis in Python: GeoPandas, spatial operations, geocoding, rasterio, interactive maps (Folium/Plotly)",
     "End-to-end spatial Python pipeline, interactive maps, raster analysis"),
    ("Module 7", "Spatial R (sf + spdep + spatialreg)", 6, 12, "Module 2 + Module 5",
     "Spatial analysis in R: sf, tmap, spdep (Moran's I, LISA), spatialreg (spatial lag/error), reproducible spatial reports",
     "Spatial regression with direct/indirect effects, LISA maps, R Markdown spatial report"),
    ("Module 8", "Interpretation & Analytical Thinking", 10, 20, "Modules 1-4 (run alongside or after)",
     "The interpretive layer: substantive vs statistical significance, reading regression output, causal claims, spatial pattern interpretation, uncertainty communication, writing results sections, critical paper reading, visualisation critique, policy translation",
     "Can interpret any quantitative output substantively; writes convincing results and discussion sections; critically reads empirical papers; translates findings for non-technical audiences"),
    ("TOTAL", "Complete Curriculum", 83, 166, "",
     "From zero to comprehensive: CS foundations, programming, statistics, econometrics, spatial analysis, and interpretive mastery",
     "Thesis-ready skills across quantitative, spatial, and communicative dimensions of social science research"),
]


# ══════════════════════════════════════════════════════════════════════════════
#  WORKBOOK BUILDER
# ══════════════════════════════════════════════════════════════════════════════

wb = openpyxl.Workbook()

# ── helper: auto-width ────────────────────────────────────────────────────────
def auto_width(ws, widths):
    for col_letter, width in widths.items():
        ws.column_dimensions[col_letter].width = width


def set_row_height(ws, row_num, height):
    ws.row_dimensions[row_num].height = height


# ── OVERVIEW SHEET ────────────────────────────────────────────────────────────
ov = wb.active
ov.title = "📋 OVERVIEW"
ov.sheet_properties.tabColor = CLR["tab_over"]

# Title
ov.merge_cells('A1:H1')
title_cell = ov['A1']
title_cell.value = "QUANTITATIVE & SPATIAL SKILLS CURRICULUM FOR SOCIAL SCIENCES"
title_cell.font = Font(name="Calibri", bold=True, size=16, color="FFFFFF")
title_cell.fill = hdr_fill("1F3864")
title_cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
ov.row_dimensions[1].height = 35

ov.merge_cells('A2:H2')
sub_cell = ov['A2']
sub_cell.value = "2 hours/day · 65 days · 130 total hours | Assumes zero prior knowledge | Designed for social sciences research (urban, development, environmental)"
sub_cell.font = Font(name="Calibri", italic=True, size=11, color="FFFFFF")
sub_cell.fill = hdr_fill("2E75B6")
sub_cell.alignment = Alignment(horizontal="center", vertical="center")
ov.row_dimensions[2].height = 22

ov_headers = ["Module", "Name", "Days", "Hours", "Prerequisites",
              "Description", "Outcomes", "Tab Colour"]
write_row(ov, 3, ov_headers, is_header=True)
ov.row_dimensions[3].height = 20

tab_colors = ["7030A0", "4472C4", "70AD47", "ED7D31", "FF0000", "9E480E", "00B0F0", "375623", "FF7F27", "1F3864"]
for i, (row_data) in enumerate(OVERVIEW_DATA):
    row_num = i + 4
    is_alt = i % 2 == 0
    write_row(ov, row_num, list(row_data) + [""], is_alt=is_alt)
    # colour the tab cell
    colour_cell = ov.cell(row=row_num, column=8)
    colour_cell.fill = PatternFill("solid", fgColor=tab_colors[i])
    colour_cell.value = "  "
    set_row_height(ov, row_num, 60)

auto_width(ov, {'A': 12, 'B': 28, 'C': 8, 'D': 8, 'E': 25, 'F': 55, 'G': 45, 'H': 12})

# Note row
note_row = len(OVERVIEW_DATA) + 4 + 2
ov.merge_cells(f'A{note_row}:H{note_row}')
note_cell = ov.cell(row=note_row, column=1)
note_cell.value = ("⚑  HOW TO USE THIS CURRICULUM: Work through modules sequentially. Each day = 2 hours. "
                   "All datasets are free and openly licensed. All software is free and open-source. "
                   "Assignments should be attempted BEFORE reading the answer/solution. "
                   "Papers marked 'Read' should be read on the same day — they are short and directly relevant. "
                   "If stuck on code: (1) re-read the error message carefully; (2) search the exact error on Stack Overflow; (3) check the package documentation.")
note_cell.font = Font(name="Calibri", italic=True, size=10, color="595959")
note_cell.fill = PatternFill("solid", fgColor="FFF2CC")
note_cell.alignment = Alignment(wrap_text=True, horizontal="left", vertical="top")
ov.row_dimensions[note_row].height = 50


# ── GENERIC SHEET BUILDER ─────────────────────────────────────────────────────
SHEET_COLS = [
    ("Day", 6),
    ("Week", 6),
    ("Module", 22),
    ("Topic", 28),
    ("Sub-Topics / Skills", 45),
    ("What to Cover & Key Concepts", 60),
    ("Primary Resource (with link)", 50),
    ("Additional Resources", 55),
    ("Dataset / Practice Problem (with link)", 55),
    ("Assignment", 60),
    ("Paper to Read", 50),
    ("Key Concepts to Master", 40),
]

def build_sheet(wb, title, tab_color, accent_hex, data):
    ws = wb.create_sheet(title=title)
    ws.sheet_properties.tabColor = tab_color
    ws.freeze_panes = 'A4'

    # Sheet header
    ws.merge_cells(f'A1:{get_column_letter(len(SHEET_COLS))}1')
    sh = ws['A1']
    sh.value = (title.replace("🐍 ", "").replace("📊 ", "").replace("🗺 ", "")
                    .replace("📈 ", "").replace("🌍 ", "").replace("🟦 ", "")
                    .replace("🟢 ", "").replace("💻 ", "").replace("🧠 ", ""))
    sh.font = Font(name="Calibri", bold=True, size=14, color="FFFFFF")
    sh.fill = hdr_fill(accent_hex)
    sh.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 28

    ws.merge_cells(f'A2:{get_column_letter(len(SHEET_COLS))}2')
    sc = ws['A2']
    sc.value = "2 hours per day | Start at zero | Build to research-ready proficiency"
    sc.font = Font(name="Calibri", italic=True, size=10, color="FFFFFF")
    sc.fill = hdr_fill(accent_hex)
    sc.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[2].height = 16

    # Column headers
    col_names = [c[0] for c in SHEET_COLS]
    col_widths = {get_column_letter(i+1): c[1] for i, c in enumerate(SHEET_COLS)}
    write_row(ws, 3, col_names, is_header=True)
    ws.row_dimensions[3].height = 22

    for i, row in enumerate(data):
        row_num = i + 4
        is_alt = i % 2 == 0
        write_row(ws, row_num, list(row), is_alt=is_alt)
        set_row_height(ws, row_num, 100)

    auto_width(ws, col_widths)

    # Module group colouring (accent border on module change)
    prev_module = None
    for i, row in enumerate(data):
        row_num = i + 4
        if row[2] != prev_module:
            for col in range(1, len(SHEET_COLS) + 1):
                c = ws.cell(row=row_num, column=col)
                c.border = Border(
                    left=Side(style='medium', color=accent_hex),
                    right=THIN,
                    top=Side(style='medium', color=accent_hex),
                    bottom=THIN
                )
            prev_module = row[2]

    return ws


# ── BUILD ALL SUBJECT SHEETS ──────────────────────────────────────────────────
build_sheet(wb, "💻 CS FOUNDATIONS", "7030A0", "7030A0", CS_DATA)
build_sheet(wb, "🐍 PYTHON", CLR["tab_py"], CLR["accent1"], PYTHON_DATA)
build_sheet(wb, "📊 R LANGUAGE", CLR["tab_r"], CLR["accent2"], R_DATA)
build_sheet(wb, "📈 STATISTICS & MODELLING", CLR["tab_stats"], "C55A11", STATS_DATA)
build_sheet(wb, "💡 ADVANCED ECONOMETRICS", CLR["tab_econ"], CLR["accent5"], ECON_DATA)
build_sheet(wb, "🗺 QGIS & SPATIAL THINKING", CLR["tab_qgis"], CLR["tab_qgis"], QGIS_DATA)
build_sheet(wb, "🌍 SPATIAL PYTHON", CLR["tab_spy"], "0070C0", SPATIAL_PY_DATA)
build_sheet(wb, "🟢 SPATIAL R", CLR["tab_sr"], CLR["accent2"], SPATIAL_R_DATA)
build_sheet(wb, "🧠 INTERPRETATION", "FF7F27", "C55A11", INTERP_DATA)


# ── SAVE ──────────────────────────────────────────────────────────────────────
output_path = "/workspace/Spatial_Quant_Skills_Curriculum.xlsx"
wb.save(output_path)
print(f"✅  Workbook saved: {output_path}")
print(f"    Sheets: {[s.title for s in wb.worksheets]}")

# Count rows
total_days = sum([len(CS_DATA), len(PYTHON_DATA), len(R_DATA), len(STATS_DATA),
                  len(ECON_DATA), len(QGIS_DATA), len(SPATIAL_PY_DATA),
                  len(SPATIAL_R_DATA), len(INTERP_DATA)])
print(f"    Total curriculum days: {total_days}")
print(f"    Total study hours: {total_days * 2}")
