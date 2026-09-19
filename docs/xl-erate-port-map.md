# XLerate → Arixcel Core port map

Source repository: [omegarhovega/XLerate](https://github.com/omegarhovega/XLerate) (MIT License)

| XLerate (TypeScript) | Arixcel (C#) |
|---|---|
| `src/core/traceBuilder.ts` | `ArixcelExplorer.Core/Tracing/TraceBuilder.cs` |
| `src/core/traceGraph.ts` | `ArixcelExplorer.Core/Tracing/TraceModels.cs` |
| `src/core/tracePolicy.ts` | `ArixcelExplorer.Core/Tracing/TracePolicy.cs` |
| `src/core/traceUtils.ts` | `ArixcelExplorer.Core/Tracing/TraceUtils.cs` |
| `src/core/formulaReferences.ts` | `ArixcelExplorer.Core/Formulas/FormulaReferences.cs` |
| `src/core/formulaConsistency.ts` | `ArixcelExplorer.Core/Audit/FormulaConsistency.cs` |
| `src/core/autoColor.ts` | `ArixcelExplorer.Core/Audit/AutoColor.cs` |
| `src/taskpane/traceDialog.ts` | `ArixcelExplorer.UI/Windows/ExplorerWindow.xaml.cs` |
| `src/taskpane/traceDialog.html` | `ArixcelExplorer.UI/Windows/ExplorerWindow.xaml` |

New Arixcel-only modules (no XLerate equivalent):

- `FormulaAstParser.cs` — logical formula tree
- `FormulaEvaluator.cs` — partial sub-expression evaluation
- `CalculationFlow.cs` — input/calc/output classification
- `CompareEngine.cs` — worksheet diff
