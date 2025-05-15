using System.Text;

public class MultiTextWriter : TextWriter
{
    private readonly TextWriter[] writers;

    public MultiTextWriter(params TextWriter[] writers)
    {
        this.writers = writers;
    }

    public override Encoding Encoding => Encoding.UTF8;

    public override void Write(char value)
    {
        foreach (var writer in writers)
        {
            writer.Write(value);
        }
    }

    public override void WriteLine(string? value)
    {
        foreach (var writer in writers)
        {
            writer.WriteLine(value);
        }
    }
    public override void Write(string? value)
    {
        foreach (var writer in writers)
        {
            writer.Write(value);
        }
    }

    public override void Write(char[] buffer, int index, int count)
    {
        foreach (var writer in writers)
        {
            writer.Write(buffer, index, count);
        }
    }

    public override void Flush()
    {
        foreach (var writer in writers)
        {
            writer.Flush();
        }
    }
}
