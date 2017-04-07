#!/usr/bin/env python
# -*- coding: utf-8 -*-
#
#    felo.py - Main GUI part of the Felo program
#
#    Copyright © 2006 Torsten Bronger <bronger@physik.rwth-aachen.de>
#
#    This file is part of the Felo program.
#
#    Felo is free software; you can redistribute it and/or modify it under
#    the terms of the MIT licence:
#
#    Permission is hereby granted, free of charge, to any person obtaining a
#    copy of this software and associated documentation files (the "Software"),
#    to deal in the Software without restriction, including without limitation
#    the rights to use, copy, modify, merge, publish, distribute, sublicense,
#    and/or sell copies of the Software, and to permit persons to whom the
#    Software is furnished to do so, subject to the following conditions:
#
#    The above copyright notice and this permission notice shall be included in
#    all copies or substantial portions of the Software.
#
#    THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
#    IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
#    FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL
#    THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
#    LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING
#    FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER
#    DEALINGS IN THE SOFTWARE.
#

__version__ = "$Revision: 214 $"
# $HeadURL: https://svn.sourceforge.net/svnroot/felo/src/felo.py $
distribution_version = "1.0"

import re, os, codecs, sys, time, StringIO, textwrap, platform, webbrowser, shutil
datapath = os.path.abspath(os.path.dirname(sys.argv[0]))
import gettext, locale
locale.setlocale(locale.LC_ALL, '')
if os.name == 'nt':
    # For Windows: set, if needed, a value in LANG environmental variable
    lang = os.getenv('LANG')
    if lang is None:
        lang = locale.getdefaultlocale()[0]  # en_US, fr_FR, el_GR etc..
    if lang:
        os.environ['LANG'] = lang
    gettext.install('felo', datapath + "/po", unicode=True)
else:
    gettext.install('felo', unicode=True)
import felo_rating
import wx, wx.grid, wx.py.editor, wx.py.editwindow, wx.html, wx.lib.hyperlink

if os.name == 'nt':
    try:
        os.chdir(os.path.expanduser('~'))
    except OSError:
        pass

class HtmlPreviewFrame(wx.Frame):
    def __init__(self, parent, title, file):
        wx.Frame.__init__(self, parent, wx.ID_ANY, title, size=(400, 600))
        self.SetIcon(App.icon)
        html = wx.html.HtmlWindow(self)
        if "gtk2" in wx.PlatformInfo:
            html.SetStandardFonts()
        html.LoadPage(file)

class Editor(wx.py.editor.EditWindow):
    def __init__(self, parent):
        wx.stc.StyledTextCtrl.__init__(self, parent, wx.ID_ANY)
        self.SetTabWidth(8)
        self.setStyles(wx.py.editwindow.FACES)
        self.SetMarginWidth(0, self.TextWidth(wx.stc.STC_STYLE_LINENUMBER, "00000"))
        self.SetMarginWidth(1, 10)
        self.SetScrollWidth(self.TextWidth(wx.stc.STC_STYLE_DEFAULT, 68*"0"))
        self.SetLexer(wx.stc.STC_LEX_CONTAINER)
        self.Bind(wx.stc.EVT_STC_STYLENEEDED, self.OnStyling)
        self.bout_line_pattern = re.compile(u"\\s*(?:(?P<date>(?:\\d{4}-\\d{1,2}-\\d{1,2})?(?:\\.?\\d+)?)"
                                            u"\\s*\t+\\s*)?"
                                            u"(?P<first>.+?)\\s*--\\s*(?P<second>.+?)\\s*\t+\\s*"
                                            u"(?P<score>\\d+:\\d+\\s*(?P<fenced_to>(?:/\\d+)|\\*)?)\\s*\\Z")
        self.item_line_pattern = re.compile(u"\\s*(?P<name>[^\t]+?)\\s*\t+\\s*(?P<value>.+?)\\s*\\Z")
    def OnStyling(self, event):
        def apply_style(span, style):
            start, end = span
            self.StartStyling(position+start, 0xff)
            self.SetStyling(end-start, style)
        start, end = self.PositionFromLine(self.LineFromPosition(self.GetEndStyled())), event.GetPosition()
        text = self.GetTextRange(start, end).encode("utf-8")
        position = start
        lines = text.splitlines(True)
        for line in lines:
            if line.lstrip().startswith("#"):
                apply_style((0, len(line)), 6)
            else:
                match = self.bout_line_pattern.match(line)
                if match:
                    apply_style(match.span(1), 3)
                    apply_style(match.span(2), 5)
                    apply_style(match.span(3), 5)
                    apply_style(match.span(5), 1)
                else:
                    match = self.item_line_pattern.match(line)
                    if match:
                        apply_style(match.span(1), 5)
            position += len(line)

class ResultFrame(wx.Frame):
    def __init__(self, title, fencerlist, *args, **keyw):
        wx.Frame.__init__(self, None, wx.ID_ANY, size=(300, 500), title=title, *args, **keyw)
        self.SetIcon(App.icon)
        grid = wx.FlexGridSizer(2, 1)
        grid.AddGrowableRow(0, 1)
        grid.AddGrowableCol(0, 1)
        result_html = u"<table><tbody>"
        self.clipboard_contents = u""
        for fencer in fencerlist:
            result_html += u"<tr><td>%s</td><td>%s</td></tr>\n" % (fencer.name, unicode(fencer.felo_rating))
            self.clipboard_contents += u"%s\t%s" % (fencer.name, unicode(fencer.felo_rating)) + os.linesep
        result_html += u"</tbody></table>"
        html = wx.html.HtmlWindow(self)
        if "gtk2" in wx.PlatformInfo:
            html.SetStandardFonts()
        html.SetPage(result_html)
        grid.Add(html, flag=wx.EXPAND)
        button = wx.Button(self, wx.ID_OK, _("Copy to clipboard"))
        self.Bind(wx.EVT_BUTTON, self.OnCopyClipboard, button)
        grid.Add(button, flag=wx.ALIGN_CENTER_HORIZONTAL)
        self.SetSizer(grid)
    def OnCopyClipboard(self, event):
        text = wx.TextDataObject(self.clipboard_contents)
        if wx.TheClipboard.Open():
#            wx.TheClipboard.UsePrimarySelection()
            wx.TheClipboard.SetData(text)
            wx.TheClipboard.Close()
            wx.TheClipboard.Flush()
        
class HTMLDialog(wx.Dialog):
    def __init__(self, directory, *args, **keyw):
        wx.Dialog.__init__(self, None, wx.ID_ANY, title=_(u"HTML export"), *args, **keyw)
        if "gtk2" in wx.PlatformInfo:
            self.SetIcon(App.icon)
        vbox_main = wx.BoxSizer(wx.VERTICAL)
        text = wx.StaticText(self, wx.ID_ANY,
                             textwrap.fill(_(u"The web files will be written to the folder \"%s\".")
                                           % directory, 41))
        vbox_main.Add(text, flag=wx.LEFT | wx.RIGHT | wx.TOP, border=10)
        vbox_checkboxes = wx.BoxSizer(wx.VERTICAL)
        self.plot_switch = wx.CheckBox(self, wx.ID_ANY, _(u"with plot"))
        vbox_checkboxes.Add(self.plot_switch, flag=wx.BOTTOM, border=10)
        self.HTML_preview = wx.CheckBox(self, wx.ID_ANY, _(u"preview the HTML"))
        self.HTML_preview.SetValue(True)
        vbox_checkboxes.Add(self.HTML_preview)
        vbox_main.Add(vbox_checkboxes, flag=wx.ALL, border=20)
        hbox_buttons = wx.BoxSizer(wx.HORIZONTAL)
        ok_button = wx.Button(self, wx.ID_OK, _("Okay"))
        ok_button.SetDefault()
        hbox_buttons.Add(ok_button)
        cancel_button = wx.Button(self, wx.ID_CANCEL, _("Cancel"))
        hbox_buttons.Add(cancel_button, flag=wx.LEFT, border=10)
        vbox_main.Add(hbox_buttons, flag=wx.ALIGN_CENTER)
        hbox_top = wx.BoxSizer(wx.HORIZONTAL)
        hbox_top.Add(vbox_main, flag=wx.ALL | wx.ALIGN_CENTER, border=5)
        self.SetSizer(hbox_top)
        self.Fit()

class AboutWindow(wx.Dialog):
    def __init__(self, *args, **keyw):
        wx.Dialog.__init__(self, None, wx.ID_ANY, title=_(u"About Felo"), *args, **keyw)
        if "gtk2" in wx.PlatformInfo:
            self.SetIcon(App.icon)
        vbox_main = wx.BoxSizer(wx.VERTICAL)
        text1 = wx.StaticText(self, wx.ID_ANY, _(u"version ")+distribution_version+
                              _(u", revision ")+__version__[11:-2])
        vbox_main.Add(text1, flag=wx.ALIGN_CENTER)
        logo = wx.StaticBitmap(self, wx.ID_ANY,
                               wx.BitmapFromImage(wx.Image(datapath+"/felo-logo-small.png", wx.BITMAP_TYPE_PNG)))
        vbox_main.Add(logo, flag=wx.ALIGN_CENTER)
        text2 = wx.StaticText(self, wx.ID_ANY, u"— "+_(u"Estimate fencing strengths of sport fencers")+u" —")
        vbox_main.Add(text2, flag=wx.ALIGN_CENTER | wx.LEFT | wx.RIGHT | wx.TOP, border=10)
        text3 = wx.StaticText(self, wx.ID_ANY, _(u"Brought to you by the sport fencing group at the\n"
                                                 u"University of Technology Aachen (RWTH), Germany."),
                              style=wx.ALIGN_CENTER)
        vbox_main.Add(text3, flag=wx.ALIGN_CENTER | wx.LEFT | wx.RIGHT | wx.TOP, border=10)
        text4 = wx.StaticText(self, wx.ID_ANY, _(u"For bug reports, suggestions, documentation, the\n"
                                                 u"source code, and the mailinglist visit Felo's homepage at"),
                              style=wx.ALIGN_CENTER)
        vbox_main.Add(text4, flag=wx.ALIGN_CENTER | wx.LEFT | wx.RIGHT | wx.TOP, border=10)
        vbox_main.Add(wx.lib.hyperlink.HyperLinkCtrl(self, wx.ID_ANY, "http://felo.sourceforge.net",
                                                     style=wx.ALIGN_CENTER),
                      flag=wx.ALIGN_CENTER | wx.LEFT | wx.RIGHT, border=10)
        text5 = wx.StaticText(self, wx.ID_ANY, _(u"English translation by Torsten Bronger."), style=wx.ALIGN_CENTER)
        vbox_main.Add(text5, flag=wx.ALIGN_CENTER | wx.LEFT | wx.RIGHT | wx.TOP, border=10)
        text6 = wx.StaticText(self, wx.ID_ANY, u"© 2006 Torsten Bronger")
        vbox_main.Add(text6, flag=wx.ALIGN_LEFT | wx.LEFT | wx.RIGHT | wx.TOP, border=10)
        hbox_top = wx.BoxSizer(wx.HORIZONTAL)
        hbox_top.Add(vbox_main, flag=wx.ALL | wx.ALIGN_CENTER, border=5)
        self.SetSizer(hbox_top)
        self.Fit()
        for widget in (self, logo, text1, text2, text3, text4, text5, text6):
            widget.Bind(wx.EVT_LEFT_DOWN, self.OnClick)
    def OnClick(self, event):
        self.Destroy()

class ProgressFrame(wx.Frame):
    def __init__(self, message, title, *args, **keyw):
        wx.Frame.__init__(self, None, wx.ID_ANY, title=title, *args, **keyw)
        self.SetIcon(App.icon)
        panel = wx.Panel(self, wx.ID_ANY)
        vbox_main = wx.BoxSizer(wx.VERTICAL)
        text = wx.StaticText(panel, wx.ID_ANY, message)
        vbox_main.Add(text, flag=wx.BOTTOM, border = 20)
        self.gauge = wx.Gauge(panel, wx.ID_ANY, 100, size=(250, 20))
        self.gauge.SetBezelFace(3)
        self.gauge.SetShadowWidth(3)
        vbox_main.Add(self.gauge, flag=wx.ALIGN_CENTER)
        hbox_top = wx.BoxSizer(wx.HORIZONTAL)
        hbox_top.Add(vbox_main, flag=wx.ALL | wx.ALIGN_CENTER, border=25)
        panel.SetSizer(hbox_top)
        hbox_top.Fit(self)
        hbox_top.SetSizeHints(self)
    def update(self, ratio):
        self.gauge.SetValue(int(round(ratio*100)))
        wx.Yield()

class Frame(wx.Frame):
    def __init__(self, *args, **keyw):
        wx.Frame.__init__(self, None, wx.ID_ANY, size=(600, 600), title="Felo", *args, **keyw)
        self.SetIcon(App.icon)

        menu_bar = wx.MenuBar()

        menu_file = wx.Menu()
        self.Bind(wx.EVT_MENU, self.OnNew, menu_file.Append(wx.ID_ANY, _(u"&New")))
        self.Bind(wx.EVT_MENU, self.OnOpen, menu_file.Append(wx.ID_ANY, _(u"&Open")))
        self.Bind(wx.EVT_MENU, self.OnSave, menu_file.Append(wx.ID_ANY, _(u"&Save")))
        self.Bind(wx.EVT_MENU, self.OnSaveAs, menu_file.Append(wx.ID_ANY, _(u"Save &as")))
        menu_file.AppendSeparator()
        self.Bind(wx.EVT_MENU, self.OnExit, menu_file.Append(wx.ID_ANY, _(u"&Quit")))
        menu_bar.Append(menu_file, _(u"&File"))

        self.editor = Editor(self)
        self.editor.SetEOLMode(wx.stc.STC_EOL_LF)

        menu_edit = wx.Menu()
        self.Bind(wx.EVT_MENU, lambda e: self.editor.Undo(), menu_edit.Append(wx.ID_ANY, _(u"&Undo")+"\tCtrl-Z"))
        self.Bind(wx.EVT_MENU, lambda e: self.editor.Redo(), menu_edit.Append(wx.ID_ANY, _(u"&Redo")+"\tCtrl-R"))
        menu_edit.AppendSeparator()
        self.Bind(wx.EVT_MENU, lambda e: self.editor.Cut(), menu_edit.Append(wx.ID_ANY, _(u"Cu&t")+"\tCtrl-X"))
        self.Bind(wx.EVT_MENU, lambda e: self.editor.Copy(), menu_edit.Append(wx.ID_ANY, _(u"&Copy")+"\tCtrl-C"))
        self.Bind(wx.EVT_MENU, lambda e: self.editor.Paste(), menu_edit.Append(wx.ID_ANY, _(u"&Paste")+"\tCtrl-V"))
        self.Bind(wx.EVT_MENU, lambda e: self.editor.Clear(), menu_edit.Append(wx.ID_ANY, _(u"Cle&ar")+"\tDEL"))
        self.Bind(wx.EVT_MENU, lambda e: self.editor.SelectAll(),
                  menu_edit.Append(wx.ID_ANY, _(u"Select a&ll")+"\tCtrl-A"))
        menu_bar.Append(menu_edit, _(u"&Edit"))

        menu_calculate = wx.Menu()
        self.Bind(wx.EVT_MENU, self.OnCalculateFeloRatings, menu_calculate.Append(wx.ID_ANY, _(u"Calculate &Felo ratings")))
        self.Bind(wx.EVT_MENU, self.OnGenerateHTML, menu_calculate.Append(wx.ID_ANY, _(u"Generate &HTML")))
        self.Bind(wx.EVT_MENU, self.OnBootstrapping, menu_calculate.Append(wx.ID_ANY, _(u"&Bootstrapping")))
        self.Bind(wx.EVT_MENU, self.OnEstimateFreshmen, menu_calculate.Append(wx.ID_ANY, _(u"&Estimate freshmen")))
        menu_bar.Append(menu_calculate, _(u"&Calculate"))

        menu_help = wx.Menu()
        self.Bind(wx.EVT_MENU, self.OnWebHelp, menu_help.Append(wx.ID_ANY, _(u"Online &help")))
        self.Bind(wx.EVT_MENU, self.OnReportBug, menu_help.Append(wx.ID_ANY, _(u"Report &bug")))
        self.Bind(wx.EVT_MENU, self.OnFeloForum, menu_help.Append(wx.ID_ANY, _(u"Felo &forum")))
        self.Bind(wx.EVT_MENU, self.OnWebpage, menu_help.Append(wx.ID_ANY, _(u"Felo &webpage")))
        menu_help.AppendSeparator()
        self.Bind(wx.EVT_MENU, self.OnShowLicence, menu_help.Append(wx.ID_ANY, _(u"Show &licence")))
        self.Bind(wx.EVT_MENU, self.OnAbout, menu_help.Append(wx.ID_ANY, _(u"&About Felo")))
        menu_bar.Append(menu_help, _(u"&Help"))

        self.SetMenuBar(menu_bar)

        self.Bind(wx.EVT_CLOSE, self.OnCloseWindow)

        self.felo_filename = _("unnamed.felo")
        self.SetTitle(u"Felo – "+os.path.split(self.felo_filename)[1])
        self.editor.Bind(wx.stc.EVT_STC_CHANGE, self.OnChange)
        self.felo_file_changed = False
        self.SendSizeEvent()
        if len(sys.argv) > 1:
            self.open_felo_file(sys.argv[1])
        self.editor.SetFocus()
    def OnWebHelp(self, event):
        webbrowser.open(_("http://felo.sourceforge.net/felo/"))
    def OnShowLicence(self, event):
        licence_window = HtmlPreviewFrame(self, _("Software licence"), datapath+"/"+_("licence.html"))
        licence_window.Show()
    def OnReportBug(self, event):
        webbrowser.open("http://sourceforge.net/tracker/?func=add&group_id=183431&atid=905214")
    def OnFeloForum(self, event):
        webbrowser.open("http://sourceforge.net/forum/forum.php?forum_id=638727")
    def OnWebpage(self, event):
        webbrowser.open("http://felo.sourceforge.net")
    def OnChange(self, event):
        self.felo_file_changed = True
    def AssureSave(self):
        if self.felo_file_changed:
            answer = wx.MessageBox(_(u"The file \"%s\" has changed.  Save it?") % self.felo_filename,
                                   _(u"File changed"), wx.YES_NO | wx.CANCEL | wx.YES_DEFAULT |
                                   wx.ICON_QUESTION, self)
            if answer == wx.YES:
                return self.OnSave(None)
            elif answer == wx.CANCEL:
                return wx.ID_CANCEL
    def read_utf8_file(self, filename):
        file = codecs.open(filename, encoding="utf-8")
        filecontents = "\n".join([line.rstrip('\r\n') for line in file]) + "\n"
        file.close()
        return filecontents
    def open_felo_file(self, felo_filename):
        felo_filename = os.path.abspath(felo_filename)
        path = os.path.dirname(felo_filename)
        try:
            os.chdir(path)
        except OSError:
            wx.MessageBox(_(u"Could not open file because folder '%s' doesn't exist.") % path,
                          _(u"Folder not found"), wx.OK | wx.ICON_ERROR, self)
            return
        self.felo_filename = felo_filename
        if os.path.isfile(self.felo_filename):
            self.editor.SetText(self.read_utf8_file(self.felo_filename))
        self.felo_file_changed = False
        self.SetTitle(u"Felo – "+os.path.split(self.felo_filename)[1])
    def OnNew(self, event):
        if self.AssureSave() == wx.ID_CANCEL:
            return
        self.felo_filename = _("unnamed.felo")
        self.editor.ClearAll()
        self.SetTitle(u"Felo – "+os.path.split(self.felo_filename)[1])
        self.editor.SetText(self.read_utf8_file(datapath+"/"+_("boilerplate.felo")))
    def OnOpen(self, event):
        if self.AssureSave() == wx.ID_CANCEL:
            return
        wildcard = _(u"Felo file (*.felo)|*.felo|"
                     "All files (*.*)|*.*")
        dialog = wx.FileDialog(None, _(u"Select Felo file"), os.getcwd(),
                               "", wildcard, wx.OPEN | wx.FILE_MUST_EXIST)
        if dialog.ShowModal() == wx.ID_OK:
            self.open_felo_file(dialog.GetPath())
            dialog.Destroy()
    def OnExit(self, event):
        self.Close()
    def OnCloseWindow(self, event):
        if self.AssureSave() == wx.ID_CANCEL:
            return
        self.Destroy()
    def save_felo_file(self):
        if self.felo_filename:
            if os.path.isfile(self.felo_filename):
                shutil.copyfile(self.felo_filename, os.path.splitext(self.felo_filename)[0]+".bak")
            file = codecs.open(self.felo_filename, "wb", encoding="utf-8")
            file.write(self.editor.GetText())
            file.close()
            self.felo_file_changed = False
            return True
        return False
    def OnSave(self, event):
        if self.felo_filename == _("unnamed.felo"):
            return self.OnSaveAs(event)
        else:
            self.save_felo_file()
            return wx.ID_OK
    def OnSaveAs(self, event):
        wildcard = _(u"Felo file (*.felo)|*.felo|"
                     "All files (*.*)|*.*")
        dialog = wx.FileDialog(None, _(u"Select Felo file"), os.getcwd(),
                               self.felo_filename, wildcard, wx.SAVE | wx.OVERWRITE_PROMPT)
        result = dialog.ShowModal()
        if result == wx.ID_OK:
            self.felo_filename = dialog.GetPath()
            os.chdir(os.path.dirname(self.felo_filename))
            self.SetTitle(u"Felo – "+os.path.split(self.felo_filename)[1])
            self.save_felo_file()
        dialog.Destroy()
        return result
    def parse_editor_contents(self):
        felo_file_contents = StringIO.StringIO(self.editor.GetText())
        felo_file_contents.name = self.felo_filename
        try:
            parameters, __, fencers, bouts = felo_rating.parse_felo_file(felo_file_contents)
        except felo_rating.LineError, e:
            self.editor.GotoLine(e.linenumber - 1)
            wx.MessageBox(_(u"Error in line %d: %s") % (e.linenumber, e.naked_description),
                          _(u"Parsing error"), wx.OK | wx.ICON_ERROR, self)
        except felo_rating.FeloFormatError, e:
            wx.MessageBox(e.description, _(u"Parsing error"), wx.OK | wx.ICON_ERROR, self)
        else:
            return parameters, fencers, bouts
        return {}, {}, []
    def report_empty_bouts(self):
        wx.MessageBox(_(u"I haven't found any bouts.  Please enter or open a\n"
                        u"complete Felo file with fencers and bouts."),
                      _(u"No bouts found"), wx.OK | wx.ICON_WARNING, self)
    def OnCalculateFeloRatings(self, event):
        parameters, fencers, bouts = self.parse_editor_contents()
        if not parameters:
            return
        if not bouts:
            self.report_empty_bouts()
            return
        fencerlist = felo_rating.calculate_felo_ratings(parameters, fencers, bouts)
        result_frame = ResultFrame(_(u"Felo ratings ") + parameters["groupname"], fencerlist)
        result_frame.Show()
    def OnGenerateHTML(self, event):
        parameters, fencers, bouts = self.parse_editor_contents()
        if not parameters:
            return
        if not bouts:
            self.report_empty_bouts()
            return
        bouts.sort()
        last_date = time.strftime(_(u"%x"), time.strptime(bouts[-1].date_string[:10], "%Y-%m-%d"))
        base_filename = parameters["groupname"].lower()
        html_dialog = HTMLDialog(parameters["output folder"])
        result = html_dialog.ShowModal()
        make_plot = html_dialog.plot_switch.GetValue()
        HTML_preview = html_dialog.HTML_preview.GetValue()
        html_dialog.Destroy()
        if result != wx.ID_OK:
            return
        file_list = u""
        try:
            html_filename = os.path.join(parameters["output folder"], base_filename+".html")
            html_file = codecs.open(html_filename, "w", "utf-8")
            file_list += base_filename+".html\n"
            print>>html_file, u"""<?xml version="1.0" encoding="utf-8"?>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head><title>%(title)s</title><meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<meta http-equiv="Content-Style-Type" content="text/css" /><style type="text/css">
/*<![CDATA[*/
@import "felo.css";
/*]]>*/
</style></head><body>\n\n<h1>%(title)s</h1>\n<h2>%(date)s</h2>\n\n<table><tbody>""" % \
                {"title": _(u"Felo ratings ")+parameters["groupname"], "date": _(u"as of ")+last_date}
            fencerlist = felo_rating.calculate_felo_ratings(parameters, fencers, bouts, plot=make_plot)
        except felo_rating.ExternalProgramError, e:
            wx.MessageBox(e.description, _(u"External program not found"), wx.OK | wx.ICON_ERROR, self)
            return
        except IOError:
            wx.MessageBox(_(u"Could not write to file or directory.\n(I tried to write to '%s'.)")
                            % parameters["output folder"], _(u"Couldn't write file"),
                          wx.OK | wx.ICON_ERROR, self)
            return
        for fencer in fencerlist:
            print>>html_file, u"<tr><td class='name'>%s</td><td class='felo-rating'>%d</td></tr>" % \
                (fencer.name, fencer.felo_rating)
        print>>html_file, u"</tbody></table>"
        if make_plot:
            print>>html_file, u"<p class='felo-plot'><img class='felo-plot' src='%s.png' alt='%s' /></p>" % \
                (base_filename, _(u"Felo ratings plot for ")+parameters["groupname"])
            file_list += base_filename+".png\n"
            if os.path.isfile(os.path.join(parameters["output folder"], base_filename+".pdf")):
                print>>html_file, _(u"<p class='printable-notice'>Also in a <a href='%s.pdf'>"
                                    u"printable version</a>.</p>") % base_filename
                file_list += base_filename+".pdf\n"
        print>>html_file, u"</body></html>"
        html_file.close()
        if HTML_preview:
            html_window = HtmlPreviewFrame(self, _(u"Preview HTML"), html_filename)
            html_window.Show()
        wx.MessageBox(_(u"The following files must be uploaded from\n%s\nto the web server:\n\n") %
                      parameters["output folder"] + file_list,
                      _(u"Upload file list"), wx.OK | wx.ICON_INFORMATION, self)
    def OnBootstrapping(self, event):
        parameters, fencers, bouts = self.parse_editor_contents()
        if not parameters:
            return
        if not bouts:
            self.report_empty_bouts()
            return
        answer = wx.MessageBox(_(u"The bootstrapping will change the fencer data.  "
                                 "Are you sure that you wish to continue?"),
                               _(u"Bootstrapping"), wx.YES_NO | wx.NO_DEFAULT |
                               wx.ICON_QUESTION, self)
        if answer == wx.NO:
            return
        progress_window = ProgressFrame(_(u"I'm bootstrapping, please be patient") + u" …", _(u"Bootstrapping"))
        progress_window.Show()
        wx.Yield()
        try:
            felo_rating.calculate_felo_ratings(parameters, fencers, bouts, bootstrapping=True,
                                               bootstrapping_callback = progress_window.update)
            for fencer in fencers.values():
                if not fencer.freshman:
                    fencer.initial_felo_rating = fencer.felo_rating
            self.editor.SetText(felo_rating.write_back_fencers(self.editor.GetText(), fencers))
            self.felo_file_changed = True
        except felo_rating.BootstrappingError, e:
            progress_window.Destroy()
            wx.MessageBox(_(u"The bootstrapping didn't converge.  You may want to increase\n"
                            u"the value for the 'threshold bootstrapping' parameter."),
                          _(u"Bootstrapping didn't converge"), wx.OK | wx.ICON_ERROR, self)
        progress_window.Destroy()
    def OnEstimateFreshmen(self, event):
        parameters, fencers, bouts = self.parse_editor_contents()
        if not parameters:
            return
        if not bouts:
            self.report_empty_bouts()
            return
        answer = wx.MessageBox(_(u"Estimating the freshmen will change their initial Felo numbers.  "
                                 u"Are you sure that you wish to continue?"),
                               _(u"Estimating freshmen"), wx.YES_NO | wx.NO_DEFAULT |
                               wx.ICON_QUESTION, self)
        if answer == wx.NO:
            return
        felo_rating.calculate_felo_ratings(parameters, fencers, bouts, estimate_freshmen=True)
        for fencer in fencers.values():
            if fencer.freshman:
                fencer.initial_felo_rating = fencer.felo_rating
        self.editor.SetText(felo_rating.write_back_fencers(self.editor.GetText(), fencers))
        self.felo_file_changed = True
    def OnAbout(self, event):
        about_window = AboutWindow()
        about_window.ShowModal()
        about_window.Destroy()

class App(wx.App):
    def OnInit(self):
        App.icon = wx.EmptyIcon()
        App.icon.CopyFromBitmap(wx.BitmapFromImage(wx.Image(datapath+"/felo-icon.png", wx.BITMAP_TYPE_PNG)))
        self.frame = Frame()
        self.frame.Show()
        self.SetTopWindow(self.frame)
        return True

app = App()
app.MainLoop()
