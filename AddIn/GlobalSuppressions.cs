﻿//-----------------------------------------------------------------------
// <copyright file="GlobalSuppressions.cs" company="Microsoft Corporation">
//     Copyright (c) Microsoft Corporation 2011. All rights reserved.
// </copyright>
//-----------------------------------------------------------------------

[assembly: System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Design", "CA2210:AssembliesShouldHaveValidStrongNames", Justification = "Assemblies will be signed later")]
[assembly: System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Performance", "CA1811:AvoidUncalledPrivateCode", Scope = "member", Target = "Microsoft.Research.Wwt.Excel.Addin.ThisAddIn.#IsCached(System.String)", Justification = "Autogenerated code")]
[assembly: System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Performance", "CA1811:AvoidUncalledPrivateCode", Scope = "member", Target = "Microsoft.Research.Wwt.Excel.Addin.ThisAddIn.#NeedsFill(System.String)", Justification = "Autogenerated code")]
[assembly: System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Performance", "CA1811:AvoidUncalledPrivateCode", Scope = "member", Target = "Microsoft.Research.Wwt.Excel.Addin.ThisAddIn.#StartCaching(System.String)", Justification = "Autogenerated code")]
[assembly: System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Performance", "CA1811:AvoidUncalledPrivateCode", Scope = "member", Target = "Microsoft.Research.Wwt.Excel.Addin.ThisAddIn.#StopCaching(System.String)", Justification = "Autogenerated code")]
[assembly: System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Globalization", "CA1300:SpecifyMessageBoxOptions", Scope = "member", Target = "Microsoft.Research.Wwt.Excel.Addin.Ribbon.#ShowError(System.String)", Justification = "Setting the option is moving the focus to active desktop")]
[assembly: System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Globalization", "CA1300:SpecifyMessageBoxOptions", Scope = "member", Target = "Microsoft.Research.Wwt.Excel.Addin.WorkflowController.#ValidateMappedColumns()", Justification = "Setting the option is moving the focus to active desktop")]
[assembly: System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Globalization", "CA1300:SpecifyMessageBoxOptions", Scope = "member", Target = "Microsoft.Research.Wwt.Excel.Addin.Ribbon.#ShowWarning(System.String)", Justification = "Setting the option is moving the focus to active desktop")]
[assembly: System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Globalization", "CA1300:SpecifyMessageBoxOptions", Scope = "member", Target = "Microsoft.Research.Wwt.Excel.Addin.WorkflowController.#ShowRangeAffectedLocalWWTWarning()", Justification = "Setting the option is moving the focus to active desktop")]
[assembly: System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Globalization", "CA1300:SpecifyMessageBoxOptions", Scope = "member", Target = "Microsoft.Research.Wwt.Excel.Addin.WorkflowController.#ShowRangeAffectedWWTWarning()", Justification = "Setting the option is moving the focus to active desktop")]
[assembly: System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Globalization", "CA1300:SpecifyMessageBoxOptions", Scope = "member", Target = "Microsoft.Research.Wwt.Excel.Addin.WorkflowController.#ShowRangeFormulaWarning()", Justification = "Setting the option is moving the focus to active desktop")]
[assembly: System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Performance", "CA1811:AvoidUncalledPrivateCode", Scope = "member", Target = "Microsoft.Research.Wwt.Excel.Addin.LayerMap.#Name", Justification = "Used in Binding with the UI.")]
[assembly: System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Maintainability", "CA1506:AvoidExcessiveClassCoupling", Scope = "type", Target = "Microsoft.Research.Wwt.Excel.Addin.WorkflowController")]
[assembly: System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Globalization", "CA1300:SpecifyMessageBoxOptions", Scope = "member", Target = "Microsoft.Research.Wwt.Excel.Addin.Ribbon.#ShowWarningWithResult(System.String)", Justification = "Setting the option is moving the focus to active desktop")]
[assembly: System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Maintainability", "CA1502:AvoidExcessiveComplexity", Scope = "member", Target = "Microsoft.Research.Wwt.Excel.Addin.LayerMap.#SetAutoMap()", Justification = "Auto mapping logic will be easy to understand if it is not broken")]
[assembly: System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Reliability", "CA2000:Dispose objects before losing scope", Scope = "member", Target = "Microsoft.Research.Wwt.Excel.Addin.PushPin.#GetPushPinBitmap(System.Int32)", Justification = "Bitmap image has to be returned and also caller method disposes Bitmap object.")]
[assembly: System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Maintainability", "CA1502:AvoidExcessiveComplexity", Scope = "member", Target = "Microsoft.Research.Wwt.Excel.Addin.LayerMap.#SetMappedColumnType()", Justification = "The entire logic has to be in a single method, else it would not be readable")]
[assembly: System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Design", "CA1031:DoNotCatchGeneralExceptionTypes", Scope = "member", Target = "Microsoft.Research.Wwt.Excel.Addin.LayerMap.#OnPropertyChangeNotifierDoWork(System.Object,System.ComponentModel.DoWorkEventArgs)", Justification = "Any exception in background thread will disable the Add-in tab")]
[assembly: System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Design", "CA1031:DoNotCatchGeneralExceptionTypes", Scope = "member", Target = "Microsoft.Research.Wwt.Excel.Addin.LayerMap.#OnPropertyChangeNotifierCompleted(System.Object,System.ComponentModel.RunWorkerCompletedEventArgs)", Justification = "Any exception in background thread will disable the Add-in tab")]
[assembly: System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Design", "CA1001:TypesThatOwnDisposableFieldsShouldBeDisposable", Scope = "type", Target = "Microsoft.Research.Wwt.Excel.Addin.LayerMap", Justification = "CancellationTokenSource is disposed properly in StopNotifying method.")]
