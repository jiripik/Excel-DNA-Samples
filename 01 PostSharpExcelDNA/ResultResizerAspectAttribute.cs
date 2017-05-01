namespace PostSharpExcelDNA
{
    using ExcelDna.Integration;
    using PostSharp.Aspects;

    using System;

    [Serializable]
    internal class ResultResizerAspectAttribute : OnMethodBoundaryAspect
    {
        private const int ExcelFormulaMaxLength = 255;

        public sealed override void OnEntry(MethodExecutionArgs args)
        {
            var excelReference = XlCall.Excel(XlCall.xlfCaller) as ExcelReference;

            var excelFormula = GetFormula(excelReference);

            if (string.IsNullOrWhiteSpace(excelFormula))
            {
                args.ReturnValue = ExcelError.ExcelErrorNA;
                args.FlowBehavior = FlowBehavior.Return;
                return;
            }

            if (excelFormula.Length > ExcelFormulaMaxLength)
            {
                args.ReturnValue = "Formula too long - Excel supports formulas up to 255 characters";
                args.FlowBehavior = FlowBehavior.Return;
                return;
            }
        }

        public sealed override void OnSuccess(MethodExecutionArgs args)
        {
            if (args.ReturnValue != null && args.ReturnValue.Equals(ExcelError.ExcelErrorNA))
            {
                args.ReturnValue = ExcelError.ExcelErrorGettingData;
                return;
            }

            if (args.ReturnValue is string)
            {
                return;
            }

            ExcelResultResizer.ResizeAndAutoFormat(args.ReturnValue as object[,]);
        }

        public sealed override void OnException(MethodExecutionArgs args)
        {
            args.ReturnValue = args.Exception.Message;
            args.FlowBehavior = FlowBehavior.Return;
        }

        private static string GetFormula(ExcelReference excelReference)
        {
            try
            {
                return (string)XlCall.Excel(XlCall.xlfGetCell, 41, excelReference);
            }
            catch
            {
                return string.Empty;
            }
        }
    }
}