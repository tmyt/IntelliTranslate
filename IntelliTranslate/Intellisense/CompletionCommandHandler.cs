using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using Microsoft.VisualStudio;
using Microsoft.VisualStudio.Language.Intellisense;
using Microsoft.VisualStudio.OLE.Interop;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Text;
using Microsoft.VisualStudio.Text.Editor;
using Microsoft.VisualStudio.TextManager.Interop;

namespace IntelliTranslate.Intellisense
{
    class CompletionCommandHandler : IOleCommandTarget
    {
        private IOleCommandTarget m_nextCommandHandler;
        private ITextView m_textView;
        private CompletionHandlerProvider m_provider;
        private ICompletionSession m_session;

        internal CompletionCommandHandler(IVsTextView textViewAdapter, ITextView textView, CompletionHandlerProvider provider)
        {
            this.m_textView = textView;
            this.m_provider = provider;

            //add the command to the command chain
            textViewAdapter.AddCommandFilter(this, out m_nextCommandHandler);
        }

        private char GetTypeChar(IntPtr pvaIn)
        {
            return (char)(ushort)Marshal.GetObjectForNativeVariant(pvaIn);
        }

        public int Exec(ref Guid pguidCmdGroup, uint nCmdID, uint nCmdexecopt, IntPtr pvaIn, IntPtr pvaOut)
        {
            bool handled = false;
            int hresult = VSConstants.S_OK;

            // 1. Pre-process
            if (pguidCmdGroup == VSConstants.VSStd2K)
            {
                switch ((VSConstants.VSStd2KCmdID)nCmdID)
                {
                    case VSConstants.VSStd2KCmdID.AUTOCOMPLETE:
                    case VSConstants.VSStd2KCmdID.COMPLETEWORD:
                        if(IsSuggestionNeeded())
                            handled = StartSession();
                        break;
                    case VSConstants.VSStd2KCmdID.RETURN:
                        handled = Complete(false);
                        break;
                    case VSConstants.VSStd2KCmdID.TAB:
                        handled = Complete(true);
                        break;
                    case VSConstants.VSStd2KCmdID.CANCEL:
                        handled = Cancel();
                        break;
                }
            }

            if (!handled)
                hresult = m_nextCommandHandler.Exec(pguidCmdGroup, nCmdID, nCmdexecopt, pvaIn, pvaOut);

            if (ErrorHandler.Succeeded(hresult))
            {
                if (pguidCmdGroup == VSConstants.VSStd2K)
                {
                    switch ((VSConstants.VSStd2KCmdID)nCmdID)
                    {
                        case VSConstants.VSStd2KCmdID.TYPECHAR:
                            char ch = GetTypeChar(pvaIn);
                            if (ch == ' ')
                                StartSession();
                            else if (m_session != null)
                                Filter();
                            break;
                        case VSConstants.VSStd2KCmdID.BACKSPACE:
                            Filter();
                            break;
                    }
                }
            }

            return hresult;
        }

        private void Filter()
        {
            if (m_session == null)
                return;

            m_session.SelectedCompletionSet.SelectBestMatch();
            m_session.SelectedCompletionSet.Recalculate();
        }

        bool Cancel()
        {
            if (m_session == null)
                return false;

            m_session.Dismiss();

            return true;
        }

        bool Complete(bool force)
        {
            if (m_session == null)
                return false;

            if (!m_session.SelectedCompletionSet.SelectionStatus.IsSelected && !force)
            {
                m_session.Dismiss();
                return false;
            }
            else
            {
                m_session.Commit();
                return true;
            }
        }

        bool StartSession()
        {
            if (m_session != null)
            {
                m_session.Dismiss();
                m_session = null;
            }

            SnapshotPoint caret = m_textView.Caret.Position.BufferPosition;
            ITextSnapshot snapshot = caret.Snapshot;

            if (!m_provider.CompletionBroker.IsCompletionActive(m_textView))
            {
                m_session = m_provider.CompletionBroker.CreateCompletionSession(m_textView, snapshot.CreateTrackingPoint(caret, PointTrackingMode.Positive), true);
            }
            else
            {
                m_session = m_provider.CompletionBroker.GetSessions(m_textView)[0];
            }
            m_session.Dismissed += (sender, args) => m_session = null;

            m_session.Start();

            return true;
        }

        private bool IsSuggestionNeeded()
        {
            var snapshot = m_textView.TextSnapshot;
            var position = m_textView.Caret.Position.BufferPosition;
            var line = position.GetContainingLine();
            var p = position.Position - line.Start.Position;
            var s = line.GetText();
            var head = s.Substring(0, p)
                .Reverse()
                .TakeWhile(c => c > 0xff && !char.IsControl(c) && !char.IsPunctuation(c) && !char.IsSeparator(c) && !char.IsSymbol(c) && !char.IsWhiteSpace(c))
                .Count();
            var tail = s.Substring(p)
                .TakeWhile(c => c > 0xff && !char.IsControl(c) && !char.IsPunctuation(c) && !char.IsSeparator(c) && !char.IsSymbol(c) && !char.IsWhiteSpace(c))
                .Count();
            var length = head + tail;
            return length != 0;
        }

        public int QueryStatus(ref Guid pguidCmdGroup, uint cCmds, OLECMD[] prgCmds, IntPtr pCmdText)
        {
            if (pguidCmdGroup == VSConstants.VSStd2K)
            {
                switch ((VSConstants.VSStd2KCmdID)prgCmds[0].cmdID)
                {
                    case VSConstants.VSStd2KCmdID.AUTOCOMPLETE:
                    case VSConstants.VSStd2KCmdID.COMPLETEWORD:
                        prgCmds[0].cmdf = (uint)OLECMDF.OLECMDF_ENABLED | (uint)OLECMDF.OLECMDF_SUPPORTED;
                        return VSConstants.S_OK;
                }
            }
            return m_nextCommandHandler.QueryStatus(pguidCmdGroup, cCmds, prgCmds, pCmdText);
        }
    }
}
