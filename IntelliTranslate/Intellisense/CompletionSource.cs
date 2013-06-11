using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.ComponentModel.Composition;
using System.Threading;
using EnvDTE;
using Microsoft.VisualStudio.Language.Intellisense;
using Microsoft.VisualStudio.Text;
using Microsoft.VisualStudio.Text.Operations;
using Microsoft.VisualStudio.Utilities;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

namespace IntelliTranslate.Intellisense
{
    class CompletionSource : ICompletionSource
    {
        private CompletionSourceProvider m_sourceProvider;
        private ITextBuffer m_textBuffer;

        public CompletionSource(CompletionSourceProvider sourceProvider, ITextBuffer textBuffer)
        {
            m_sourceProvider = sourceProvider;
            m_textBuffer = textBuffer;
        }

        void ICompletionSource.AugmentCompletionSession(ICompletionSession session, IList<CompletionSet> completionSets)
        {
            var completions = new List<KeyValuePair<string, string>>();

            var snapshot = session.TextView.TextSnapshot;
            var position = session.GetTriggerPoint(snapshot);
            if (position.HasValue)
            {
                var line = position.Value.GetContainingLine();
                var p = position.Value.Position - line.Start.Position;
                var s = line.GetText();
                var head = s.Substring(0, p)
                    .Reverse()
                    .TakeWhile(c => c > 0xff && !char.IsControl(c) && !char.IsPunctuation(c) && !char.IsSeparator(c) && !char.IsSymbol(c) && !char.IsWhiteSpace(c))
                    .Count();
                var tail = s.Substring(p)
                    .TakeWhile(c => c > 0xff && !char.IsControl(c) && !char.IsPunctuation(c) && !char.IsSeparator(c) && !char.IsSymbol(c) && !char.IsWhiteSpace(c))
                    .Count();
                var length = head + tail;
                var t = s.Substring(p - head, length);
                if (!string.IsNullOrWhiteSpace(t))
                {
                    var dte = (DTE)m_sourceProvider.ServiceProvider.GetService(typeof(DTE));
                    dte.StatusBar.Animate(true, vsStatusAnimation.vsStatusAnimationFind);
                    dte.StatusBar.Progress(true, "IntelliTranslate: 翻訳中...", 0, 1000);
                    // Googleにクエリかける
                    var url = "http://translate.google.com/translate_a/t?client=t&hl=ja&sl=ja&tl=en&ie=UTF-8&oe=UTF-8&multires=1&oc=2&otf=1&ssel=6&tsel=3&sc=1&q=";
                    url += WebUtility.UrlEncode(t);
                    var req = WebRequest.CreateHttp(url);
                    req.UserAgent = "User-Agent: Opera/9.80 (Windows NT 6.2; WOW64) Presto/2.12.388 Version/12.15";
                    req.Referer = "http://translate.google.com/";
                    try
                    {
                        var resp = req.GetResponse();
                        using (var stream = resp.GetResponseStream())
                        using (var reader = new StreamReader(stream))
                        {
                            var ary = JsonConvert.DeserializeObject<JArray>(reader.ReadToEnd());
                            var result = new KeyValuePair<string, string>(ary[0][0][0].Value<string>(), ary[0][0][1].Value<string>());
                            var compls = ary[1]
                                .SelectMany(token => token[2])
                                .Select(token => new KeyValuePair<string, string>(token[0].Value<string>(), string.Join(", ", token[1])))
                                .ToArray();
                            completions.Add(result);
                            completions.AddRange(compls);
                        }
                        resp.Close();
                    }
                    catch(WebException e)
                    {
                        
                        dte.StatusBar.Text = "IntelliTranslate: エラー " + e.Message;
                    }catch
                    {
                    }
                    dte.StatusBar.Progress(false);
                    dte.StatusBar.Animate(false, vsStatusAnimation.vsStatusAnimationFind);
                }
                if (completions.Count > 0)
                {
                    completionSets.Add(new CompletionSet(
                                           "Translation", //the non-localized title of the tab 
                                           "翻訳", //the display title of the tab
                                           snapshot.CreateTrackingSpan(position.Value.Position - head, length, SpanTrackingMode.EdgeInclusive),
                                           completions.Select(kv => new Completion(kv.Key, kv.Key, kv.Value, null, null)),
                                           null));
                }
            }
        }

        private ITrackingSpan FindTokenSpanAtPosition(ITrackingPoint point, ICompletionSession session)
        {
            SnapshotPoint currentPoint = (session.TextView.Caret.Position.BufferPosition) - 1;
            ITextStructureNavigator navigator = m_sourceProvider.NavigatorService.GetTextStructureNavigator(m_textBuffer);
            TextExtent extent = navigator.GetExtentOfWord(currentPoint);
            return currentPoint.Snapshot.CreateTrackingSpan(extent.Span, SpanTrackingMode.EdgeInclusive);
        }

        private bool m_isDisposed;
        public void Dispose()
        {
            if (!m_isDisposed)
            {
                GC.SuppressFinalize(this);
                m_isDisposed = true;
            }
        }
    }
}
