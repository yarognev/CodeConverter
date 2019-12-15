using System.Linq;
using System.Threading.Tasks;
using Xunit;

namespace CodeConverter.Tests
{
    /// <summary>
    /// TODO:
    /// * Ensure input VB has no error diagnostics
    /// * Add repros for:
    ///   *  MyBase instead of forms when using static MyForms property
    ///   *  Double qualification (related to ambiguous references perhaps)
    ///   *  Empty "global::" generic for anonymous types
    /// * Reduce amount of repro code
    /// </summary>
    public class VbSimplificationServiceBugRepros
    {
        [Fact]
        public async Task EmptyArgListImmediatelyAfterFunctionRemoved()
        {
            var validInputVb = @"Friend Class TestClass
    Private prop As String
    Private prop2 As String
    Private ReadOnly Property [Property] As String
        Get
            Return If(prop, Function()
                                prop2 = CreateProperty
                                Return prop2
                            End Function()) '< This empty arg list gets removed sometimes
        End Get
    End Property
    Private Function CreateProperty() As String
        Return """"
    End Function
End Class";

            var simplified = await VbReduceHelpers.ReduceVbAsync(validInputVb);
            Assert.True(simplified.Contains("End Function()"), "Input:\r\n" + validInputVb + "\r\n\r\nOutput:\r\n" + simplified);
        }
    }
}
