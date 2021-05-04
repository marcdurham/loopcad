using LoopCAD.WPF;
using Xunit;

namespace LoopCAD.Tests
{
    public class SnakeCaseTests
    {
        [Fact]
        public void Normal()
        {
            Assert.Equal(
                "this_is_pascal",
                SnakeCase.Convert("ThisIsPascal"));
        }
    }
}
