"""
Step 2: Multi-Tool — Explore Before Executing

Step 1 had a single tool (execute_python). The agent had to guess the data
structure. Now we add two exploration tools:

  1. list_available_datasets() — see what files exist
  2. read_dataset(filename, n_rows) — preview first N rows

The agent's workflow becomes:
  Question → list datasets → read relevant ones → write code → execute

This is MULTI-HOP tool calling: the agent calls tools SEQUENTIALLY,
where each call informs the next decision.

Watch the logs: the agent will call list_available_datasets first, THEN
decide which file to read, THEN write code using the actual column names.

Test queries:
  - 

Dependencies:
    pip install openai e2b-code-interpreter python-dotenv

Run:
    python step2_multi_tool.py
"""

import os
import json
from openai import OpenAI
from e2b_code_interpreter import Sandbox
from dotenv import load_dotenv
import pandas as pd
import PyPDF2
from docx import Document


load_dotenv()

MODEL = "gpt-4o-mini"
client = OpenAI(api_key=os.getenv("OPENAI_API_KEY"))

MTECH_DIR = r"D:\Kaustav\MTECH\Admission"
MAX_FILE_SIZE_MB = 50


def discover_files():
    """Recursively discover supported files in the MTech directory."""

    supported_extensions = {
        ".csv",
        ".xlsx",
        ".xls",
        ".json",
        ".txt",
        ".docx",
        ".pdf"
    }

    files = {}

    for root, dirs, filenames in os.walk(MTECH_DIR):

        for filename in filenames:

            extension = os.path.splitext(filename)[1].lower()

            if extension in supported_extensions:

                filepath = os.path.join(root, filename)

                # Preserve folder structure relative to MTECH_DIR
                relative_path = os.path.relpath(
                    filepath,
                    MTECH_DIR
                )

                # Normalize Windows \ to /
                relative_path = relative_path.replace("\\", "/")

                files[relative_path] = filepath

    return files


FILES = discover_files()




# ─────────────────────────────────────────────
# E2B SANDBOX SETUP
# ─────────────────────────────────────────────

def create_sandbox():
    """Create E2B sandbox and upload discovered MTech files."""

    sandbox = Sandbox.create(timeout=450)

    for relative_path, filepath in FILES.items():

        file_size_mb = os.path.getsize(filepath) / (1024 * 1024)

        if file_size_mb > MAX_FILE_SIZE_MB:
            print(
                f"  [SKIP] {relative_path} "
                f"({file_size_mb:.1f} MB) exceeds "
                f"{MAX_FILE_SIZE_MB} MB limit"
            )
            continue

        sandbox_path = f"/home/user/{relative_path}"

        directory = os.path.dirname(sandbox_path)

        if directory:
            sandbox.files.make_dir(directory)

        print(f"  [UPLOAD] {relative_path} ({file_size_mb:.1f} MB)")

        with open(filepath, "rb") as f:
            sandbox.files.write(
                sandbox_path,
                f.read()
            )

    return sandbox




# ─────────────────────────────────────────────
# TOOL IMPLEMENTATIONS
# ─────────────────────────────────────────────






def list_available_datasets():
    """List all supported files available in the MTech directory."""

    if not FILES:
        return "No supported files were found in the MTech directory."

    output = ["Available files:"]

    for filename, filepath in FILES.items():
        extension = os.path.splitext(filename)[1].lower()

        file_type_map = {
            ".csv": "CSV",
            ".xlsx": "Excel",
            ".xls": "Excel",
            ".json": "JSON",
            ".txt": "Text",
            ".docx": "Word",
            ".pdf": "PDF"
        }

        file_type = file_type_map.get(extension, "Unknown")

        try:
            file_size = os.path.getsize(filepath)

            if file_size < 1024:
                size = f"{file_size} B"
            elif file_size < 1024 ** 2:
                size = f"{file_size / 1024:.1f} KB"
            else:
                size = f"{file_size / (1024 ** 2):.1f} MB"

        except OSError:
            size = "Unknown"

        output.append(
            f"- {filename} | Type: {file_type} | Size: {size}"
        )

    return "\n".join(output)




def read_dataset(filename, n_rows=5):
    """Preview a file from the MTech directory based on its file type."""

    if filename not in FILES:
        return f"ERROR: File '{filename}' not found in MTech directory."

    filepath = FILES[filename]
    extension = os.path.splitext(filename)[1].lower()

    try:
        # CSV
        if extension == ".csv":
            import pandas as pd

            df = pd.read_csv(filepath, nrows=n_rows)
            return df.to_string(index=False)

        # Excel
        elif extension in [".xlsx", ".xls"]:
            import pandas as pd

            excel_file = pd.ExcelFile(filepath)

            output = []
            output.append(
                f"Excel file: {filename}"
            )
            output.append(
                f"Sheets: {', '.join(excel_file.sheet_names)}"
            )

            for sheet_name in excel_file.sheet_names:
                df = pd.read_excel(
                    filepath,
                    sheet_name=sheet_name,
                    nrows=n_rows
                )

                output.append(
                    f"\n--- Sheet: {sheet_name} ---"
                )
                output.append(
                    df.to_string(index=False)
                )

            return "\n".join(output)

        # JSON
        elif extension == ".json":
            import pandas as pd

            try:
                df = pd.read_json(filepath)

                return df.head(n_rows).to_string(index=False)

            except ValueError:
                with open(filepath, "r", encoding="utf-8") as f:
                    data = json.load(f)

                if isinstance(data, list):
                    preview = data[:n_rows]
                else:
                    preview = data

                return json.dumps(
                    preview,
                    indent=2,
                    ensure_ascii=False
                )

        # TXT
        elif extension == ".txt":
            with open(filepath, "r", encoding="utf-8") as f:
                lines = f.readlines()

            return "".join(lines[:n_rows + 1])

        # PDF
        elif extension == ".pdf":
            import PyPDF2

            output = []

            with open(filepath, "rb") as f:
                reader = PyPDF2.PdfReader(f)

                pages_to_read = min(n_rows, len(reader.pages))

                output.append(
                    f"PDF file: {filename}"
                )
                output.append(
                    f"Total pages: {len(reader.pages)}"
                )

                for i in range(pages_to_read):
                    text = reader.pages[i].extract_text()

                    output.append(
                        f"\n--- Page {i + 1} ---"
                    )
                    output.append(
                        text if text else "(No extractable text)"
                    )

            return "\n".join(output)

        # DOCX
        elif extension == ".docx":
            document = Document(filepath)

            output = []
            output.append(
                f"Word document: {filename}"
            )

            paragraphs = []

            for paragraph in document.paragraphs:
                text = paragraph.text.strip()

                if text:
                    paragraphs.append(text)

            preview = paragraphs[:n_rows]

            output.append(
                f"Total paragraphs: {len(paragraphs)}"
            )

            for i, text in enumerate(preview, 1):
                output.append(
                    f"\n--- Paragraph {i} ---"
                )
                output.append(text)

            # Include basic table information
            if document.tables:
                output.append(
                    f"\nTotal tables: {len(document.tables)}"
                )

                for table_index, table in enumerate(
                    document.tables[:3], 1
                ):
                    output.append(
                        f"\n--- Table {table_index} ---"
                    )

                    for row in table.rows[:n_rows]:
                        row_data = [
                            cell.text.strip()
                            for cell in row.cells
                        ]

                        output.append(
                            " | ".join(row_data)
                        )

            return "\n".join(output)
        

        # Unsupported format
        else:
            return (
                f"ERROR: File type '{extension}' is not supported "
                f"for preview."
            )

    except Exception as e:
        return (
            f"ERROR: Unable to preview '{filename}': "
            f"{type(e).__name__}: {str(e)}"
        )


    


def preflight_check(code):
    """Compile the code locally BEFORE spending a sandbox round-trip.

    Returns an error string (sent back to the LLM), or None if the code parses.
    A local compile() also lets us diagnose one defect the model reliably
    CANNOT see on its own: when it double-escapes newlines in its tool-call
    JSON, the code arrives as one long line full of literal backslash-n pairs.
    The model can't distinguish that from real newlines in its own context,
    so a plain traceback never breaks the loop — it needs to be told exactly
    what happened at the encoding level.
    """
    try:
        compile(code, "<agent_code>", "exec")
        return None
    except SyntaxError as e:
        if "\\n" in code and "\n" not in code:
            return (
                "ERROR: PreflightSyntaxError: your code arrived as a SINGLE line "
                "containing literal backslash-n ('\\n') two-character sequences "
                "instead of real newline characters. This is a JSON string-escaping "
                "mistake in your tool call, NOT a logic bug — do not change the "
                "analysis logic. Resend the SAME code, but with actual newlines "
                "between statements (encode line breaks as \\n exactly once in the "
                "JSON string, not \\\\n)."
            )
        return (
            f"ERROR: PreflightSyntaxError: the code failed Python compilation before "
            f"reaching the sandbox: {e.msg} (line {e.lineno}). Fix the syntax and resend."
        )


def execute_python(sandbox, code, error_analysis=None):
    """Execute Python code in the E2B sandbox."""
    preflight_error = preflight_check(code)
    if preflight_error:
        return preflight_error

    execution = sandbox.run_code(code)
    stdout = "\n".join(execution.logs.stdout).strip()

    if execution.error:
        result = f"ERROR: {execution.error.name}: {execution.error.value}\n{execution.error.traceback}"
        # CRUCIAL: never discard output printed before the crash. The model's
        # own debug prints (e.g. df.columns) are often exactly the evidence it
        # needs to fix the error — hiding them traps it in a retry loop.
        if stdout:
            result += (
                "\n\nPARTIAL OUTPUT (printed before the error — use it to debug):\n"
                + stdout
            )
        return result

    return stdout if stdout else "(no output)"


# ─────────────────────────────────────────────
# TOOL DEFINITIONS (OpenAI format)
# ─────────────────────────────────────────────

TOOLS = [
    {
        "type": "function",
        "function": {
            "name": "list_available_datasets",
            "description": "List all supported files available in the MTech directory, including filename, file type, and file size." 
                           "You can call this to understand what data is available & its schema.",
            "parameters": {
                "type": "object",
                "properties": {},
                "required": []
            }
        }
    },
    {
        "type": "function",
        "function": {
            "name": "read_dataset",
            "description": "Preview the first N rows of a file from the MTech directory. "
                           "Supported formats include CSV, Excel (.xlsx/.xls), "
                            "JSON, TXT, PDF, and Word (.docx). "
                            "Use this to understand the file structure, "
                            "contents, column names, sheets, paragraphs, "
                            "and sample data. "
                            "This is ONLY a preview and must not be used "
                            "as the full dataset for analysis. "
                            "For actual analysis, use execute_python with "
                            "the required file available in the sandbox "
                            "at '/home/user/<filename>'.",
            "parameters": {
                "type": "object",
                "properties": {
                    "filename": {
                        "type": "string",
                        "description": "Name of a file discovered in the MTech directory."
                    },
                "n_rows": {
                    "type": "integer",
                    "description": "Number of rows to preview.",
                    "default": 5
                    }
                },
                "required": ["filename"]
            }
        }
    },
    {
        "type": "function",
        "function": {
            "name": "execute_python",
            "description":  "Execute Python code in the E2B sandbox to perform analysis "
                            "on files from the MTech directory. "
                            "The available files have been uploaded to /home/user/. "
                            "Load the required files using the appropriate Python library "
                            "based on the file type. "
                            "For CSV use pandas.read_csv(), "
                            "for Excel use pandas.read_excel(), "
                            "for JSON use pandas.read_json() or json, "
                            "for text files use Python file I/O, "
                            "for PDF use an appropriate PDF library such as PyPDF2, "
                            "and for Word (.docx) files use python-docx. "
                            "Always load the full file required for the analysis; "
                            "do not use preview data. "
                            "Always use print() to output the analysis result.",
            "parameters": {
                "type": "object",
                "properties": {
                    "code": {
                        "type": "string",
                        "description": "Python code to execute. Must use print() for output."
                    },
                    "error_analysis": {
                        "type": "string",
                        "description": "Fill this ONLY when this call is a RETRY — i.e., a previous "
                                       "execute_python call in this conversation raised an error OR ran "
                                       "'successfully' but produced a wrong/empty/suspicious result. "
                                       "In 2-3 plain-English sentences, diagnose what was wrong with the "
                                       "previous code and state how this new code fixes it. "
                                       "OMIT this field entirely on a first attempt."
                    }
                },
                "required": ["code"]
            }
        }
    }
]


# ─────────────────────────────────────────────
# SYSTEM PROMPT
# ─────────────────────────────────────────────

SYSTEM_PROMPT = """
You are an MTech Copilot and a general-purpose data analysis assistant.
For every analytical question, determine the required analysis, write Python code to perform the analysis, and use the execute_python() tool to execute that Python code against the full files available in the sandbox.

You help the user analyze files available in the MTech directory.


AVAILABLE TOOLS
---------------
1. list_available_datasets()
   - Lists all supported files currently available in the MTech directory.
   - Use this first when you need to understand what files are available.

2. read_dataset(filename, n_rows)
   - Provides a preview of the selected file.
   - Supported formats:
       * CSV
       * Excel (.xlsx, .xls)
       * JSON
       * TXT
       * PDF
       * Word (.docx)
   - Use this tool to understand the structure and contents of a file before writing analysis code.

3. execute_python(code, error_analysis)
   - Executes Python code inside the E2B sandbox.
   - All discovered files are uploaded to:
       /home/user/<filename>
   - Use the appropriate Python library depending on the file type.
   - The sandbox contains the full files, not just the preview returned by read_dataset().
   - Use print() to display important analysis results.

FILE ANALYSIS WORKFLOW
----------------------
Follow this general workflow:

1. Understand the user's question.

2. Identify which file or files are relevant to the question.
   - If you do not know which files are relevant, call list_available_datasets().
   - Do not assume that a particular file exists.

3. Preview the relevant files using read_dataset() when necessary.
   - Inspect columns, sheets, structure, sample records, or document contents.
   - For Excel files, consider the relevant sheet names.
   - For PDF/TXT files, inspect the available text before deciding how to process it.

4. Before executing Python code, carefully review the code you are about to run.
   Check:
   - File names and paths
   - Column names
   - Data types
   - Join keys
   - Filtering conditions
   - Aggregations
   - Missing-value handling
   - Whether the calculation actually answers the user's question

5. Use execute_python() for the actual analysis.
   - Load the full required files from /home/user/.
   - Do NOT perform the final analysis using only the preview returned by read_dataset().
   - Use pandas for CSV and Excel analysis where appropriate.
   - Use pandas or the json library for JSON files where appropriate.
   - Use standard Python file handling for TXT files.
   - Use an appropriate PDF library for PDF processing.
   - For Word documents, use python-docx to read paragraphs and tables.

6. If multiple files are required:
   - Load each required file.
   - Verify their structures and keys.
   - Join or combine them only when the relationship is justified by the data.

7. Return the analysis result to the user in clear, understandable language.
   - Include important numbers, trends, comparisons, or findings.
   - Do not expose unnecessary implementation details unless the user asks for them.
   - Do not claim a result that was not actually obtained from the data.

SELF-REVIEW BEFORE EXECUTION
----------------------------
Before calling execute_python(), review the generated Python code.

Pay particular attention to:
- Incorrect column names
- Incorrect file paths
- Incorrect assumptions about the data
- Invalid pandas operations
- Incorrect groupby or aggregation logic
- Incorrect joins
- Division by zero
- Empty datasets
- Missing values
- Incorrect date/time handling
- Incorrect filtering
- Incorrect statistical calculations

If the code contains an obvious problem, correct it before execution.

ERROR HANDLING AND RETRY
------------------------
If execute_python() returns an error:

1. Read the error carefully.
2. Identify the likely cause.
3. Explain the cause briefly in the error_analysis field.
4. Correct the Python code.
5. Execute the corrected code again.

Do not repeat the exact same failing code unless the error indicates a temporary or external issue.

If the code executes successfully but the output is empty, suspicious, incomplete, or inconsistent with the question:

1. Review the code and assumptions.
2. Identify the likely reason.
3. Modify the code.
4. Execute it again.

GENERAL RULES
-------------
- All sandbox files are located under /home/user/.
- Use the exact discovered filename when accessing a file.
- Do not invent files, columns, sheets, values, or relationships.
- Do not assume every file is a CSV.
- Select the appropriate processing method based on the file type.
- Use previews for understanding the data and full files for actual analysis.
- Keep Python output focused on the information required to answer the user's question.
- If the available files do not contain enough information to answer the question, clearly state that.
- When appropriate, combine information from multiple files to answer the question.
- Always prioritize correctness of the analysis over speed.
"""


# ─────────────────────────────────────────────
# TOOL EXECUTOR
# ─────────────────────────────────────────────

def execute_tool(name, args, sandbox):
    if name == "list_available_datasets":
        return list_available_datasets()
    elif name == "read_dataset":
        return read_dataset(args["filename"], args.get("n_rows", 5))
    elif name == "execute_python":
        return execute_python(sandbox, args["code"], args.get("error_analysis"))
    else:
        return f"Unknown tool: {name}"


# ─────────────────────────────────────────────
# CONSOLE LOGGING
#
# One "hop" = one LLM response. If that response contains SEVERAL tool
# calls, the model requested them IN PARALLEL (it didn't need the result
# of one to ask for the next). Hops themselves are always sequential —
# that's the multi-hop part.
# ─────────────────────────────────────────────

def format_tool_call(name, args):
    """Render a tool call as a readable one-liner."""
    if name == "read_dataset":
        return f"read_dataset(\"{args['filename']}\", n_rows={args.get('n_rows', 5)})"
    if name == "execute_python":
        return "execute_python(code=<shown below>)"
    return f"{name}()"


def log_hop_header(hop, n_calls):
    if n_calls > 1:
        print(f"\n  ┌─ HOP {hop} ─ ⇉ PARALLEL: {n_calls} tool calls in ONE LLM response")
    else:
        print(f"\n  ┌─ HOP {hop} ─ → sequential: 1 tool call")


def log_tool_call(i, n, name, args):
    print(f"  │ [TOOL {i}/{n}] {format_tool_call(name, args)}")

    if name == "execute_python":
        if args.get("error_analysis"):
            print("  │")
            print("  │ [ERROR ANALYSIS — why the previous attempt failed]")
            for line in args["error_analysis"].split("\n"):
                print(f"  │   {line}")
        print("  │")
        print("  │ [CODE TO EXECUTE]")
        for line in args["code"].split("\n"):
            print(f"  │   {line}")
        print("  │")


def log_tool_result(i, n, name, args, result):
    if name == "list_available_datasets":
        print(f"  │ [RESULT {i}/{n}] {len(FILES)} files available")
    elif name == "read_dataset":
        n_rows = args.get("n_rows", 5)
        print(f"  │ [RESULT {i}/{n}] preview: first {n_rows} rows of {args['filename']}")
    else:  # execute_python — show full output
        print(f"  │ [RESULT {i}/{n}]")
        for line in result.split("\n"):
            print(f"  │   {line}")


# ─────────────────────────────────────────────
# HARNESS REMINDERS
#
# Optional tool parameters (like error_analysis) are easy for the model
# to forget: the instruction sits far away in the system prompt while the
# model is focused on the fresh traceback. So when an execution fails,
# the HARNESS appends a reminder to the tool result itself — the message
# the model reads right before writing its retry. Real coding agents
# (Claude Code, Cursor, etc.) inject feedback into tool results the
# same way.
# ─────────────────────────────────────────────

ERROR_REMINDER = (
    "\n\n[SYSTEM REMINDER] This execution FAILED. Your next execute_python call is a "
    "RETRY, so you MUST fill its error_analysis parameter: 2-3 sentences on what was "
    "wrong in this code and how the new code fixes it. Do not retry without it."
)

SUSPICIOUS_REMINDER = (
    "\n\n[SYSTEM REMINDER] This result looks suspicious (empty or no output) — often a "
    "sign of a wrong join key or filter. If it is genuinely the correct answer, proceed. "
    "If you retry with fixed code, you MUST fill the error_analysis parameter of "
    "execute_python explaining what was wrong and how you fixed it."
)


def with_harness_reminder(name, result):
    """Return (content_for_llm, reminder_label). Reminder goes to the LLM only."""
    if name != "execute_python":
        return result, None
    if result.startswith("ERROR:"):
        return result + ERROR_REMINDER, "execution failed → fill error_analysis on retry"
    if result == "(no output)" or "Empty DataFrame" in result:
        return result + SUSPICIOUS_REMINDER, "suspicious result → diagnose before retrying"
    return result, None


# ─────────────────────────────────────────────
# AGENT LOOP
# ─────────────────────────────────────────────

MAX_HOPS = 15  # runaway guard — a stuck retry loop burns tokens fast


def run_agent(user_query, sandbox):
    messages = [
        {"role": "system", "content": SYSTEM_PROMPT},
        {"role": "user", "content": user_query}
    ]

    print(f"\n  [AGENT] Thinking...")
    tool_calls_count = 0
    hop = 0

    while True:
        if hop >= MAX_HOPS:
            print(f"\n  [STATS] hop limit ({MAX_HOPS}) reached — stopping. {tool_calls_count} tool calls total")
            return ("(Agent stopped: hop limit reached without a final answer. "
                    "The model was likely stuck in a retry loop.)")

        response = client.chat.completions.create(
            model=MODEL,
            messages=messages,
            tools=TOOLS,
            max_tokens=1500,
        )

        #print("\n $$ the full responses getting from LLM is --|\n",response )

    
        choice = response.choices[0]

        #print("\n $$ And this is the choice we take --|\n",response )


        if choice.finish_reason == "tool_calls" or choice.message.tool_calls:
            messages.append(choice.message)

            hop += 1
            calls = choice.message.tool_calls
            #print("\n || the calls are\n", calls)

            n = len(calls)
            log_hop_header(hop, n)

            for i, tool_call in enumerate(calls, 1):
                name = tool_call.function.name
                args = json.loads(tool_call.function.arguments)
                tool_calls_count += 1

                log_tool_call(i, n, name, args)
                result = execute_tool(name, args, sandbox)
                log_tool_result(i, n, name, args, result)

                content, reminder = with_harness_reminder(name, result)
                if reminder:
                    print(f"  │ [HARNESS ⇒ LLM] {reminder}")

                messages.append({
                    "role": "tool",
                    "tool_call_id": tool_call.id,
                    "content": content,
                })

            print(f"  └─ HOP {hop} done")
            continue

        print(f"\n  [STATS] {hop} hop{'s' if hop != 1 else ''}, {tool_calls_count} tool calls total")
        return choice.message.content


# ─────────────────────────────────────────────
# MAIN
# ─────────────────────────────────────────────

if __name__ == "__main__":
    print("=" * 60)
    print("   STEP 2: Multi-Tool — Explore Before Executing")
    print("=" * 60)
    print("""
  Now the agent has 3 tools:
    1. list_available_datasets() — discover what data exists
    2. read_dataset(file, n) — preview columns and sample rows
    3. execute_python(code) — run analysis in sandbox

  How to read the logs:
    - Each ┌─ HOP N ─┐ box = ONE LLM response (hops run in sequence
      — that's multi-hop: each hop's results inform the next).
    - ⇉ PARALLEL inside a hop = the model requested several tool
      calls in that single response, without waiting between them.
    - [ERROR ANALYSIS] = the model diagnosed its previous failed
      attempt before retrying with fixed code.

  Type 'quit' or 'bye' or 'exit' to exit.
""")

    sandbox = create_sandbox()
    print(f"  [SANDBOX] Created + uploaded {len(FILES)} files\n")

    try:
        while True:
            query = input("  You: ").strip()
            if not query:
                continue
            if query.lower() in  ("quit", "exit", "bye"):
                break

            answer = run_agent(query, sandbox)
            print(f"\n  Agent: {answer}\n")

    finally:
        sandbox.kill()
        print("\n  [SANDBOX] Closed")
