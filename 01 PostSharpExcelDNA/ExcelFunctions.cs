namespace PostSharpExcelDNA
{
    using System;
    using ExcelDna.Integration;

    public class ExcelFunctions : IExcelAddIn 
    {
        [ExcelFunction(IsMacroType = true)]
        [ResultResizerAspect]
        public static object TryMe(object parameter)
        {
            return ExcelAsyncUtil.Run(
                "TryMe",
                new[] { parameter },
                () =>
                {
                    try
                    {
                        const int n = 100;
                        var res = new object[n, n];
                        for (var i = 0; i < n; ++i)
                        {
                            for (var j = 0; j < n; ++j)
                            {
                                res[i, j] = string.Format("{0}, {1}", i, j);
                            }
                        }

                        return res;
                    }
                    catch (Exception exception)
                    {
                        return exception.Message;
                    }
                });
        }

        public void AutoOpen()
        {
        }

        public void AutoClose()
        {
        }
    }
}