using ExcelFileSplitter.Core;
using Xunit;

namespace ExcelFileSplitter.Tests;

public class SplitExecutorTests : IDisposable
{
    private readonly string _folder;
    private readonly string _source;

    public SplitExecutorTests()
    {
        _folder = Path.Combine(Path.GetTempPath(), $"efs-exec-{Guid.NewGuid():N}");
        Directory.CreateDirectory(_folder);
        _source = Path.Combine(_folder, "Zrodlo.xlsx");
        File.WriteAllText(_source, "");
    }

    private SplitConfig Config() => new()
    {
        Mode = SplitMode.Worksheet,
        SourceWorkbook = _source,
        OutputFolder = Path.Combine(_folder, "out"),
        FileNameTemplate = "Raport_{value}.xlsx",
        Verify = true,
    };

    private static Recipient Odbiorca(string name) => new(name, [SplitKey.Of(name)], ["Dashboard"]);

    private sealed class FakeStrategy(bool verificationOk) : ISplitStrategy
    {
        public void PrepareOnce(IProgressReporter reporter) { }

        public SplitEvidence? CreateRecipientWorkbook(Recipient recipient, string outputPath, IProgressReporter reporter)
        {
            File.WriteAllText(outputPath, "dane wszystkich odbiorcow");
            return null;
        }

        public IVerificationResult Verify(
            string outputPath, Recipient recipient, SplitEvidence? evidence, IProgressReporter reporter) =>
            new FakeVerification(verificationOk);

        public void Dispose() { }
    }

    private sealed record FakeVerification(bool Ok) : IVerificationResult
    {
        public string Summary() => Ok ? "ok" : "kontrola nie przeszla";
    }

    [Fact]
    public void Plik_ktory_przeszedl_kontrole_zostaje_pod_swoja_nazwa()
    {
        SplitConfig config = Config();

        List<SplitOutcome> outcomes = SplitExecutor.CreateRecipientFiles(
            config, [Odbiorca("Anna")], new FakeStrategy(verificationOk: true), NullProgressReporter.Instance);

        SplitOutcome outcome = Assert.Single(outcomes);
        Assert.True(outcome.Ok);
        Assert.Equal(config.OutputPathFor("Anna"), outcome.Path);
        Assert.True(File.Exists(outcome.Path));
    }

    [Fact]
    public void Plik_odrzucony_dostaje_w_nazwie_NIE_WYSYLAC()
    {
        SplitConfig config = Config();

        List<SplitOutcome> outcomes = SplitExecutor.CreateRecipientFiles(
            config, [Odbiorca("Anna")], new FakeStrategy(verificationOk: false), NullProgressReporter.Instance);

        SplitOutcome outcome = Assert.Single(outcomes);
        Assert.False(outcome.Ok);
        Assert.Contains("_NIE_WYSYŁAĆ", outcome.Path);
        Assert.True(File.Exists(outcome.Path));

        Assert.False(File.Exists(config.OutputPathFor("Anna")));
    }

    [Fact]
    public void Gdy_nie_da_sie_zmienic_nazwy_odrzucony_plik_jest_kasowany()
    {
        SplitConfig config = Config();
        string outputPath = config.OutputPathFor("Anna");

        Directory.CreateDirectory(Path.GetFullPath(config.OutputFolder));
        Directory.CreateDirectory(SplitExecutor.RejectedPathFor(outputPath));

        List<SplitOutcome> outcomes = SplitExecutor.CreateRecipientFiles(
            config, [Odbiorca("Anna")], new FakeStrategy(verificationOk: false), NullProgressReporter.Instance);

        SplitOutcome outcome = Assert.Single(outcomes);
        Assert.False(outcome.Ok);

        Assert.False(File.Exists(outputPath));
        Assert.Equal("", outcome.Path);
        Assert.Contains("skasowany", outcome.FileName);
    }

    [Fact]
    public void Plik_ktorego_tworzenie_sie_wywrocilo_tez_nie_zostaje_pod_finalna_nazwa()
    {
        SplitConfig config = Config();

        List<SplitOutcome> outcomes = SplitExecutor.CreateRecipientFiles(
            config, [Odbiorca("Anna")], new FailingStrategy(), NullProgressReporter.Instance);

        SplitOutcome outcome = Assert.Single(outcomes);
        Assert.False(outcome.Ok);
        Assert.NotNull(outcome.Error);
        Assert.False(File.Exists(config.OutputPathFor("Anna")));
        Assert.Contains("_NIE_WYSYŁAĆ", outcome.Path);
    }

    [Fact]
    public void Kopia_skasowana_po_bledzie_nie_dostaje_nazwy_NIE_WYSYLAC()
    {
        SplitConfig config = Config();
        string outputPath = config.OutputPathFor("Anna");

        List<SplitOutcome> outcomes = SplitExecutor.CreateRecipientFiles(
            config, [Odbiorca("Anna")], new FailingStrategy(deletesCopy: true), NullProgressReporter.Instance);

        SplitOutcome outcome = Assert.Single(outcomes);
        Assert.False(outcome.Ok);
        Assert.Equal("", outcome.Path);
        Assert.Contains("skasowany", outcome.FileName);
        Assert.False(File.Exists(outputPath));
        Assert.False(File.Exists(SplitExecutor.RejectedPathFor(outputPath)));
    }

    private sealed class FailingStrategy(bool deletesCopy = false) : ISplitStrategy
    {
        public void PrepareOnce(IProgressReporter reporter) { }

        public SplitEvidence? CreateRecipientWorkbook(Recipient recipient, string outputPath, IProgressReporter reporter)
        {
            File.WriteAllText(outputPath, "dane wszystkich odbiorcow");
            if (deletesCopy) File.Delete(outputPath);
            throw new InvalidOperationException("Excel przestal odpowiadac");
        }

        public IVerificationResult Verify(
            string outputPath, Recipient recipient, SplitEvidence? evidence, IProgressReporter reporter) =>
            throw new InvalidOperationException("nie powinno dojsc do weryfikacji");

        public void Dispose() { }
    }

    [Fact]
    public void Przerwanie_zatrzymuje_petle_miedzy_plikami()
    {
        SplitConfig config = Config();
        using var cancelled = new CancellationTokenSource();
        cancelled.Cancel();

        List<SplitOutcome> outcomes = SplitExecutor.CreateRecipientFiles(
            config, [Odbiorca("Anna"), Odbiorca("Piotr")], new FakeStrategy(verificationOk: true),
            NullProgressReporter.Instance, cancelled.Token);

        Assert.Empty(outcomes);
    }

    public void Dispose()
    {
        try { Directory.Delete(_folder, recursive: true); } catch { }
    }

    [Fact]
    public void Podzial_tworzy_tylko_wskazanych_odbiorcow()
    {
        SplitConfig config = Config();
        Recipient[] wszyscy = [Odbiorca("Anna"), Odbiorca("Piotr"), Odbiorca("Maria")];

        List<SplitOutcome> outcomes = SplitExecutor.CreateRecipientFiles(
            config, [wszyscy[0], wszyscy[2]], new FakeStrategy(verificationOk: true), NullProgressReporter.Instance);

        Assert.Equal(2, outcomes.Count);
        Assert.True(File.Exists(config.OutputPathFor("Anna")));
        Assert.True(File.Exists(config.OutputPathFor("Maria")));
        Assert.False(File.Exists(config.OutputPathFor("Piotr")));
    }
}
