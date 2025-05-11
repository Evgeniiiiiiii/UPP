namespace ConsoleApp1
{
    public class Stream
    {
        public int Id { get; set; }
        public int DepartmentId { get; set; }
        public List<Group> Groups { get; set; } = new List<Group>();
    }
}