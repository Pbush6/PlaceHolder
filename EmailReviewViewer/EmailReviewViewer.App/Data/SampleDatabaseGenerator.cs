namespace EmailReviewViewer.App.Data;

public static class SampleDatabaseGenerator
{
    private static readonly string[] Names =
    [
        "Alice Morgan", "Bob Chen", "Carol Diaz", "David Patel", "Erin Brooks",
        "Frank Wilson", "Grace Kim", "Hector Rivera", "Ivy Thompson", "Jack Nelson"
    ];

    private static readonly string[] Subjects =
    [
        "Quarterly budget review", "Project Phoenix status", "Customer renewal follow-up",
        "Operations weekly update", "Travel approval request", "Product roadmap discussion",
        "Contract review notes", "Security training reminder", "Hiring panel feedback",
        "Invoice reconciliation"
    ];

    private static readonly string[] Folders =
    [
        @"Mailbox\Inbox", @"Mailbox\Sent Items", @"Mailbox\Projects\Phoenix",
        @"Mailbox\Finance", @"Mailbox\Archive\2025"
    ];

    public static async Task GenerateAsync(
        string databasePath,
        int count,
        IProgress<int>? progress = null,
        CancellationToken cancellationToken = default)
    {
        if (count < 1)
            throw new ArgumentOutOfRangeException(nameof(count));

        Microsoft.Data.Sqlite.SqliteConnection.ClearAllPools();
        if (File.Exists(databasePath))
            File.Delete(databasePath);
        if (File.Exists(databasePath + "-wal"))
            File.Delete(databasePath + "-wal");
        if (File.Exists(databasePath + "-shm"))
            File.Delete(databasePath + "-shm");

        var repository = new EmailRepository(databasePath);
        await repository.EnsureCreatedAsync(cancellationToken);
        var random = new Random(20260713);
        var start = new DateTime(2022, 1, 1, 8, 0, 0, DateTimeKind.Utc);

        for (var offset = 0; offset < count; offset += 1000)
        {
            var batchCount = Math.Min(1000, count - offset);
            var batch = new List<EmailMessage>(batchCount);
            for (var index = 0; index < batchCount; index++)
            {
                var number = offset + index + 1;
                var senderIndex = random.Next(Names.Length);
                var recipientIndex = (senderIndex + 1 + random.Next(Names.Length - 1)) % Names.Length;
                var received = start.AddMinutes(random.Next(0, 2_300_000));
                var subject = Subjects[number % Subjects.Length];
                var senderName = Names[senderIndex];
                var recipientName = Names[recipientIndex];
                var body = $"""
                    Hello {recipientName},

                    This is message {number:N0} regarding {subject.ToLowerInvariant()}.
                    The project team reviewed the current records, action items, schedule, and supporting documents.
                    Reference keyword {(number % 7 == 0 ? "phoenix" : "standard")} is included for repeatable search testing.

                    Regards,
                    {senderName}
                    """;
                batch.Add(new EmailMessage
                {
                    Id = number,
                    FolderPath = Folders[number % Folders.Length],
                    SenderName = senderName,
                    SenderAddress = Address(senderName),
                    ToRecipients = $"{recipientName} <{Address(recipientName)}>",
                    CcRecipients = number % 5 == 0 ? $"Records Team <records@example.com>" : "",
                    Subject = subject,
                    SentUtc = received.AddMinutes(-random.Next(1, 60)),
                    ReceivedUtc = received,
                    Preview = $"This is message {number:N0} regarding {subject.ToLowerInvariant()}.",
                    BodyText = body,
                    MessageClass = "IPM.Note",
                    EntryId = $"SAMPLE-{number:D8}"
                });
            }

            await repository.InsertBatchAsync(batch, cancellationToken);
            progress?.Report(offset + batchCount);
        }
    }

    private static string Address(string name) =>
        name.Replace(" ", ".", StringComparison.Ordinal).ToLowerInvariant() + "@example.com";
}
