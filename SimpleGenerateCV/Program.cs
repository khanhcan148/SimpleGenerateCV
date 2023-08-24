using System.Collections.Immutable;
using System.Threading.Tasks.Dataflow;
using Bogus.Extensions.Sweden;
using Xceed.Document.NET;
using Xceed.Words.NET;

BufferBlock<DocInfo> bufferBlock = new BufferBlock<DocInfo>();
ActionBlock<DocInfo> transformCreateCv = new ActionBlock<DocInfo>(async s =>
{
    DocX doc = DocX.Create(@"E:\cvTest\FakeCv\" + s.FileName);
    doc.InsertParagraph(s.Name);
    doc.InsertParagraph(s.Email);
    doc.InsertParagraph(s.Phone);
    doc.InsertParagraph(s.Address);
    doc.InsertParagraph(s.Dob.ToShortDateString());
    doc.InsertParagraph(s.Gender);
    Paragraph otherInfo = doc.InsertParagraph(s.OtherInfo, false, new Formatting
    {
        FontFamily = new Font("Century Gothic")
    });
    otherInfo.Alignment = Alignment.both;

    doc.Save();
}, new ExecutionDataflowBlockOptions() { MaxDegreeOfParallelism = 6 });

bufferBlock.LinkTo(transformCreateCv);

PersonFaker faker = new PersonFaker();
ImmutableList<DocInfo> persons = faker.GenerateLazy(5000).ToImmutableList();
foreach (DocInfo p in persons)
{
    bufferBlock.SendAsync(p);
}


Console.ReadKey();

internal class DocInfo
{
    public string Name { get; set; }
    public string Email { get; set; }
    public string Phone { get; set; }
    public string Address { get; set; }
    public DateTime Dob { get; set; }
    public string Gender { get; set; }
    public string OtherInfo { get; set; }
    public string FileName { get; set; }
}

internal class PersonFaker : Bogus.Faker<DocInfo>
{
    public PersonFaker()
    {
        RuleFor(x => x.Email, x => x.Person.Email);
        RuleFor(x => x.Name, x => x.Person.FullName);
        RuleFor(x => x.Phone, x => x.Person.Phone);
        RuleFor(x => x.Address, x => x.Person.Address.Street + ", " + x.Person.Address.City + ", " + x.Person.Address.State);
        RuleFor(x => x.Dob, x => x.Person.DateOfBirth);
        RuleFor(x => x.Gender, x => x.Person.Gender.ToString());
        RuleFor(x => x.OtherInfo, x => x.Person.Random.Words(x.Random.Number(1000, 5000)));
        RuleFor(x => x.FileName, x => x.Person.FirstName + "_" + x.Person.LastName + "_" + x.Person.Personnummer() + ".docx");
    }
}