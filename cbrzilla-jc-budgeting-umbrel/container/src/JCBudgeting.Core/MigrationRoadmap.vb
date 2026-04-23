Namespace Global.JCBudgeting.Core

    Public Class MigrationStep
        Public Property Title As String
        Public Property Description As String
        Public Property Status As String
    End Class

    Public Module MigrationRoadmap

        Public Function GetDesktopMigrationSteps() As List(Of MigrationStep)
            Return New List(Of MigrationStep) From {
                New MigrationStep With {
                    .Title = "Shared Core",
                    .Description = "Keep reusable budgeting logic in a UI-agnostic library.",
                    .Status = "Started"
                },
                New MigrationStep With {
                    .Title = "Avalonia Shell",
                    .Description = "Stand up the new cross-platform desktop app with navigation, theme, and layout primitives.",
                    .Status = "Started"
                },
                New MigrationStep With {
                    .Title = "Mobile Access",
                    .Description = "Stand up the companion server and mobile web experience so the budget can be reviewed and adjusted from another device.",
                    .Status = "Started"
                },
                New MigrationStep With {
                    .Title = "Budget Workspace",
                    .Description = "Keep refining the main budget screen with desktop-first grids, summaries, and editing workflows.",
                    .Status = "Started"
                }
            }
        End Function

    End Module

End Namespace
