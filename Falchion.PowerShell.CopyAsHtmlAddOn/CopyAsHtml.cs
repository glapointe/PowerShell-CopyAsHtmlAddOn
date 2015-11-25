using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Threading;
using Falchion.PowerShell.CopyAsHtmlAddOn.Utilities;
using Microsoft.PowerShell.Host.ISE;
using Microsoft.VisualStudio.Text;
using Microsoft.VisualStudio.Text.Editor;
using Microsoft.VisualStudio.Text.Operations;

namespace Falchion.PowerShell.CopyAsHtmlAddOn
{
    public class CopyAsHtml : IAddOnToolHostObject
    {
        #region Implementation of IAddOnToolHostObject

        public ObjectModelRoot HostObject { get; set; }

        #endregion

        public void Copy()
        {
            ISEEditor editor = null;
            ISEEditor scriptOrCommandFocusedEditor = GetScriptOrCommandFocusedEditor();
            if (scriptOrCommandFocusedEditor != null)
                editor = scriptOrCommandFocusedEditor;
            else
                editor = GetFocusedOutputControl();

            IEditorOperations editorOperations = ReflectionUtilities.GetPropertyValue(editor, "EditorOperations") as IEditorOperations;
            if (editorOperations != null)
            {
                ITextView textView = ReflectionUtilities.GetFieldValue(editorOperations, "_textView") as ITextView;
                if (textView != null)
                {
                    DispatchHelper.DispatchInvoke(((DispatcherObject) textView).Dispatcher, delegate {
                        string rtf = ReflectionUtilities.ExecuteMethod(editorOperations, "GenerateRtf",
                                                          new[] {typeof (NormalizedSnapshotSpanCollection)},
                                                          new[] {textView.Selection.SelectedSpans}) as string;
                        string html = null;
                        if (!string.IsNullOrEmpty(rtf))
                        {
                            MarkupConverter.MarkupConverter mc = new MarkupConverter.MarkupConverter();
                            html = mc.ConvertRtfToHtml(rtf);
                        }
                        CopyToClipboard(editor.SelectedText, rtf, html, textView.Selection.Mode == TextSelectionMode.Box);
                        return null;
                    });
                }
            }
        }

        private static bool CopyToClipboard(string textData, string rtfData, string htmlData, bool addBoxCutCopyTag)
        {
            try
            {
                DataObject data = new DataObject();
                
                if (rtfData != null)
                {
                    data.SetData(DataFormats.Rtf, rtfData);
                }
                if (htmlData != null)
                {
                    data.SetData(DataFormats.Html, htmlData);
                    data.SetText(htmlData);
                }
                else
                {
                    data.SetText(textData);
                }
                if (addBoxCutCopyTag)
                {
                    data.SetData("MSDEVColumnSelect", new object());
                }
                Clipboard.SetDataObject(data, true);
                return true;
            }
            catch (ExternalException)
            {
                return false;
            }
        }


        internal ISEEditor GetScriptOrCommandFocusedEditor()
        {
            PowerShellTab selectedPowerShellTab = HostObject.PowerShellTabs.SelectedPowerShellTab;
            ISEEditor lastEditorWithFocus = ReflectionUtilities.GetPropertyValue(selectedPowerShellTab, "LastEditorWithFocus") as ISEEditor;
            if (lastEditorWithFocus != null)
            {
                if (lastEditorWithFocus == selectedPowerShellTab.ConsolePane)
                {
                    return lastEditorWithFocus;
                }
                ISEFile selectedFile = selectedPowerShellTab.Files.SelectedFile;
                if ((selectedFile != null) && (selectedFile.Editor == lastEditorWithFocus))
                {
                    return lastEditorWithFocus;
                }
            }
            return null;
        }

        private ISEEditor GetFocusedOutputControl()
        {
            PowerShellTab selectedPowerShellTab = HostObject.PowerShellTabs.SelectedPowerShellTab;
            IWpfTextViewHost editorViewHost = ReflectionUtilities.GetPropertyValue(selectedPowerShellTab.ConsolePane, "EditorViewHost") as IWpfTextViewHost;
            if (editorViewHost != null && editorViewHost.TextView.VisualElement.IsFocused)
            {
                return selectedPowerShellTab.ConsolePane;
            }
            return null;
        }


    }
}
