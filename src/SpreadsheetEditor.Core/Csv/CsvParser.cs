using System.Globalization;
using System.Text;

namespace SpreadsheetEditor.Core.Csv;

/// <summary>
/// Splits decoded .csv text into records of values (RFC 4180, read leniently where that loses nothing). It reads the text once, in
/// blocks, so it never holds more of the file than the records it returns.
/// </summary>
internal sealed class CsvParser
{
    private readonly TextReader reader;
    private readonly char[] buffer = new char[16 * 1024];
    private readonly StringBuilder value = new();
    private int position;
    private int length;

    // Text read ahead by DetectSeparator, read again before the buffer.
    private string? pending;
    private int pendingPosition;

    private char separator = ',';

    // The last character Next returned, or -1 before the first.
    private int lastRead = -1;

    public CsvParser(TextReader reader) => this.reader = reader;

    /// <summary>
    /// True when the last character read is a line break: once <see cref="ReadRecord"/> has returned null, whether the text ends with one.
    /// </summary>
    public bool EndsWithLineBreak => lastRead is '\r' or '\n';

    /// <summary>The 1-based line the parser is on.</summary>
    public int LineNumber { get; private set; } = 1;

    /// <summary>
    /// Detects the separator from the header (the first record that is not a blank line) and returns it; the header is then read again by
    /// <see cref="ReadRecord"/>. The rule:
    /// <list type="bullet">
    /// <item>Count the commas, semicolons and tabs in the header outside quoted sections. The most frequent one is the separator.</item>
    /// <item>A <c>"</c> opens a quoted section only at the start of the line or right after a comma, semicolon or tab. Inside it, <c>""</c>
    /// is a quote and any other <c>"</c> closes it; line breaks inside it do not end the header.</item>
    /// <item>Ties go to comma, then semicolon, then tab. A header with none of them (a single column) is comma-separated.</item>
    /// </list>
    /// Examples: <c>a;b,c;d</c> is semicolon-separated, <c>"x,y";z</c> is semicolon-separated, <c>a,b;c</c> is comma-separated.
    /// </summary>
    public char DetectSeparator()
    {
        var raw = new StringBuilder();
        int commas = 0, semicolons = 0, tabs = 0;
        bool inQuotes = false, atValueStart = true, recordStarted = false;
        while (true)
        {
            var c = Next();
            if (c < 0)
            {
                break;
            }
            var ch = (char)c;
            raw.Append(ch);
            if (inQuotes)
            {
                if (ch == '"')
                {
                    if (Peek() == '"')
                    {
                        raw.Append((char)Next());
                    }
                    else
                    {
                        inQuotes = false;
                    }
                }
                continue;
            }
            if (ch is '\r' or '\n')
            {
                if (ch == '\r' && Peek() == '\n')
                {
                    raw.Append((char)Next());
                }
                if (recordStarted)
                {
                    break;
                }
                continue;
            }
            recordStarted = true;
            switch (ch)
            {
                case ',':
                    commas++;
                    atValueStart = true;
                    break;
                case ';':
                    semicolons++;
                    atValueStart = true;
                    break;
                case '\t':
                    tabs++;
                    atValueStart = true;
                    break;
                case '"' when atValueStart:
                    inQuotes = true;
                    atValueStart = false;
                    break;
                default:
                    atValueStart = false;
                    break;
            }
        }
        pending = raw.ToString();
        pendingPosition = 0;
        separator = semicolons > commas && semicolons >= tabs ? ';'
            : tabs > commas && tabs > semicolons ? '\t'
            : ',';
        return separator;
    }

    /// <summary>The next record's values, skipping blank lines, or null at the end of the text.</summary>
    public List<string>? ReadRecord()
    {
        while (true)
        {
            var c = Peek();
            if (c < 0)
            {
                return null;
            }
            if (c != '\r' && c != '\n')
            {
                break;
            }
            ReadLineBreak();
        }
        var values = new List<string>();
        while (true)
        {
            values.Add(ReadValue());
            var c = Peek();
            if (c == separator)
            {
                Next();
                continue;
            }
            if (c >= 0)
            {
                ReadLineBreak();
            }
            return values;
        }
    }

    private string ReadValue()
    {
        value.Clear();
        if (Peek() == '"')
        {
            var openedOnLine = LineNumber;
            Next();
            while (true)
            {
                var c = Next();
                if (c < 0)
                {
                    throw new CsvReadException(
                        CsvReadFailure.UnterminatedQuote,
                        openedOnLine,
                        string.Create(CultureInfo.InvariantCulture, $"The quoted value that begins on line {openedOnLine} is not closed."));
                }
                if (c == '"')
                {
                    if (Peek() != '"')
                    {
                        break;
                    }
                    Next();
                }
                else if (c == '\n' || (c == '\r' && Peek() != '\n'))
                {
                    LineNumber++;
                }
                Append(c);
            }
        }
        // Unquoted text, or text after a closing quote, up to the separator or the end of the line, kept as it is.
        while (true)
        {
            var c = Peek();
            if (c < 0 || c == separator || c == '\r' || c == '\n')
            {
                return value.ToString();
            }
            Next();
            Append(c);
        }
    }

    private void Append(int c)
    {
        if (c == 0)
        {
            throw new CsvReadException(
                CsvReadFailure.UnsupportedEncoding,
                LineNumber,
                "The file holds a null character: it is not UTF-8, nor UTF-16 with a byte order mark.");
        }
        value.Append((char)c);
    }

    private void ReadLineBreak()
    {
        if (Next() == '\r' && Peek() == '\n')
        {
            Next();
        }
        LineNumber++;
    }

    private int Peek()
    {
        if (pending is not null)
        {
            if (pendingPosition < pending.Length)
            {
                return pending[pendingPosition];
            }
            pending = null;
        }
        if (position == length)
        {
            length = reader.Read(buffer, 0, buffer.Length);
            position = 0;
            if (length == 0)
            {
                return -1;
            }
        }
        return buffer[position];
    }

    private int Next()
    {
        var c = Peek();
        if (c >= 0)
        {
            if (pending is not null)
            {
                pendingPosition++;
            }
            else
            {
                position++;
            }
            lastRead = c;
        }
        return c;
    }
}
