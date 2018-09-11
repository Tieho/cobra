/*
 * Cobra WinLDTP 4.0
 * 
 * Author: Nagappan Alagappan <nalagappan@vmware.com>
 * Copyright: Copyright (c) 2011-13 VMware, Inc. All Rights Reserved.
 * License: MIT license
 * 
 * http://ldtp.freedesktop.org
 * 
 * Permission is hereby granted, free of charge, to any person obtaining a copy
 * of this software and associated documentation files (the "Software"), to deal
 * in the Software without restriction, including without limitation the rights to
 * use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies
 * of the Software, and to permit persons to whom the Software is furnished to do
 * so, subject to the following conditions:
 * 
 * The above copyright notice and this permission notice shall be included in all
 * copies or substantial portions of the Software.
 * 
 * THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
 * IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
 * FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
 * AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
 * LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
 * OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
 * SOFTWARE.
*/
using System;
using System.Windows;
using System.Collections;
using CookComputing.XmlRpc;
using System.Windows.Forms;
using System.Windows.Automation;
using System.Collections.Generic;

namespace Ldtpd
{
    class Tree
    {
        Utils utils;
        static int maxPagesSelectRowSearches = 0;
        public Tree(Utils utils)
        {
            this.utils = utils;
        }
        private void LogMessage(Object o)
        {
            utils.LogMessage(o);
        }
        private AutomationElement GetObjectHandle(string windowName,
            string objName, ControlType[] type = null, bool waitForObj = true)
        {
            if (type == null)
                type = new ControlType[7] { ControlType.Tree,
                    ControlType.List, ControlType.Table,
                    ControlType.DataItem, ControlType.ListItem,
                    ControlType.TreeItem, ControlType.Custom };
            try
            {
                return utils.GetObjectHandle(windowName,
                    objName, type, waitForObj);
            }
            finally
            {
                type = null;
            }
        }
        public int DoesRowExist(String windowName, String objName,
            String text, bool partialMatch = false)
        {
            if (String.IsNullOrEmpty(text))
            {
                LogMessage("Argument cannot be empty.");
                return 0;
            }
            ControlType[] type;
            AutomationElement childHandle;
            AutomationElement elementItem;
            try
            {
                childHandle = GetObjectHandle(windowName,
                    objName, null, false);
                if (!utils.IsEnabled(childHandle))
                {
                    childHandle = null;
                    LogMessage("Object state is disabled");
                    return 0;
                }
                childHandle.SetFocus();
                type = new ControlType[4] { ControlType.TreeItem,
                    ControlType.ListItem, ControlType.DataItem,
                    ControlType.Custom };
                if (partialMatch)
                    text += "*";
                elementItem = utils.GetObjectHandle(childHandle,
                    text, type, false);
                if (elementItem != null)
                {
                    return 1;
                }
            }
            catch (Exception ex)
            {
                LogMessage(ex);
            }
            finally
            {
                type = null;
                childHandle = elementItem = null;
            }
            return 0;
        }
        private int SearchRowHelper(AutomationElement childHandle, String text,
            bool partialMatch, bool searchDown, int[] treePos)
        {
            Keyboard kb = new Keyboard(utils);

            ControlType[] type = new ControlType[4] { ControlType.TreeItem,
                ControlType.ListItem, ControlType.DataItem,
                ControlType.Custom };
            if (partialMatch)
                text += "*";

            AutomationElement rowElement = utils.GetObjectHandle(childHandle,
                text, type, false);

            // Found the element
            if(rowElement != null)
            {
                // Now let's make sure the element is really in view
                // (seems that a valid element is sometimes returned
                // even if it is under the view header)
                Rect rowRect = rowElement.Current.BoundingRectangle;
                int count = 0;

                if(searchDown)
                {
                    while(rowRect.Y >= treePos[1] + treePos[3])
                    {
                        kb.GenerateKeyEvent("<pgdown>");

                        rowElement = utils.GetObjectHandle(childHandle,
                            text, type, false);
                        rowRect = rowElement.Current.BoundingRectangle;

                        // Make sure we don't hang here if something is wrong
                        count++;

                        if(count > 100)
                            return 0;
                    }
                }
                else
                {
                    while(rowRect.Y - rowRect.Height <= treePos[1])
                    {
                        kb.GenerateKeyEvent("<pgup>");

                        rowElement = utils.GetObjectHandle(childHandle,
                            text, type, false);
                        rowRect = rowElement.Current.BoundingRectangle;

                        // Make sure we don't hang here if something is wrong
                        count++;

                        if(count > 100)
                            return 0;
                    }
                }

                return 1;
            }

            return 0;
        }
        private int SearchRowWorker(String windowName, String objName,
            String text, bool partialMatch = false, int maxPages = 10, bool searchDown = true)
        {
            AutomationElement childHandle = GetObjectHandle(windowName,
                objName);

            Generic gen = new Generic(utils);
            int[] treePos = gen.GetObjectSize(windowName, objName);

            if(SearchRowHelper(childHandle, text, partialMatch, searchDown, treePos) == 1)
                return 1;

            childHandle.SetFocus();
            utils.InternalClick(childHandle);

            Keyboard kb = new Keyboard(utils);

            // There doesn't seem to be any way to know if we've scrolled all they way up or down,
            // so user has to tell how many pages to scroll at maximum
            for(int p = 0; p < maxPages; p++)
            {
                if(searchDown)
                    kb.GenerateKeyEvent("<pgdown>");
                else
                    kb.GenerateKeyEvent("<pgup>");

                if(SearchRowHelper(childHandle, text, partialMatch, searchDown, treePos) == 1)
                    return 1;
            }

            return 0;
        }
        public int SearchRow(String windowName, String objName,
            String text, bool partialMatch = false, int maxPages = 10, bool searchDown = true)
        {
            return SearchRowWorker(windowName, objName, text, partialMatch, maxPages, searchDown);
        }
        public void SetMaxPagesSelectRowSearches(int maxPages)
        {
            maxPagesSelectRowSearches = maxPages;
        }
        public int SelectRow(String windowName, String objName,
            String text, bool partialMatch = false)
        {
            if(maxPagesSelectRowSearches > 0)
            {
                if(SearchRowWorker(windowName, objName, text, partialMatch, maxPagesSelectRowSearches, true) == 0)
                    SearchRowWorker(windowName, objName, text, partialMatch, maxPagesSelectRowSearches, false);
            }

            if (String.IsNullOrEmpty(text))
            {
                throw new XmlRpcFaultException(123, "Argument cannot be empty.");
            }
            Object pattern;
            ControlType[] type;
            AutomationElement elementItem;
            AutomationElement childHandle = GetObjectHandle(windowName,
                objName);
            if (!utils.IsEnabled(childHandle))
            {
                childHandle = null;
                throw new XmlRpcFaultException(123,
                    "Object state is disabled");
            }
            try
            {
                try
                {
                    childHandle.SetFocus();
                }
                catch (InvalidOperationException ex)
                {
                    LogMessage(ex);
                }
                if (partialMatch)
                    text = "*" + text + "*";
                type = new ControlType[4] { ControlType.TreeItem,
                    ControlType.ListItem, ControlType.DataItem,
                    ControlType.Custom };
                elementItem = utils.GetObjectHandle(childHandle,
                    text, type, true);
                if (elementItem != null)
                {
                    elementItem.SetFocus();
                    LogMessage(elementItem.Current.Name + " : " +
                        elementItem.Current.ControlType.ProgrammaticName);
                    if (elementItem.TryGetCurrentPattern(
                        SelectionItemPattern.Pattern, out pattern))
                    {
                        LogMessage("SelectionItemPattern");
                        //((SelectionItemPattern)pattern).Select();
                        // NOTE: Work around, as the above doesn't seem to work
                        // with UIAComWrapper and UIAComWrapper is required
                        // to Edit value in Spin control
                        utils.InternalClick(elementItem);
                        return 1;
                    }
                    else if (elementItem.TryGetCurrentPattern(
                        LegacyIAccessiblePattern.Pattern, out pattern))
                    {
                        utils.InternalClick(elementItem);
                        return 1;
                    }                    
                    else if (elementItem.TryGetCurrentPattern(
                        ExpandCollapsePattern.Pattern, out pattern))
                    {
                        LogMessage("ExpandCollapsePattern");
                        ((ExpandCollapsePattern)pattern).Expand();
                        return 1;
                    }
                    else
                    {
                        throw new XmlRpcFaultException(123,
                            "Unsupported pattern.");
                    }
                }
            }
            catch (Exception ex)
            {
                LogMessage(ex);
                if (ex is XmlRpcFaultException)
                    throw;
                else
                    throw new XmlRpcFaultException(123,
                        "Unhandled exception: " + ex.Message);
            }
            finally
            {
                type = null;
                pattern = null;
                elementItem = childHandle = null;
            }
            throw new XmlRpcFaultException(123,
                "Unable to find the item in list: " + text);
        }
        public int MultiSelect(String windowName, String objName,
            String[] texts, bool partialMatch = false)
        {
            if (texts == null)
            {
                throw new XmlRpcFaultException(123, "Argument cannot be empty.");
            }
            Object pattern;
            ControlType[] type;
            string[] searchTexts = texts;
            AutomationElement elementItem;
            List<string> myCollection = new List<string>();
            AutomationElement childHandle = GetObjectHandle(windowName,
                objName);
            if (!utils.IsEnabled(childHandle))
            {
                childHandle = null;
                throw new XmlRpcFaultException(123,
                    "Object state is disabled");
            }
            try
            {
                try
                {
                    childHandle.SetFocus();
                }
                catch (InvalidOperationException ex)
                {
                    LogMessage(ex);
                }
                if (partialMatch)
                {
                    foreach(string text in texts)
                        myCollection.Add("*" + text + "*");
                    // Search for the partial text, rather than the given text
                    searchTexts = myCollection.ToArray();
                }
                type = new ControlType[4] { ControlType.TreeItem,
                    ControlType.ListItem, ControlType.DataItem,
                    ControlType.Custom };
                foreach (string text in searchTexts)
                {
                    elementItem = utils.GetObjectHandle(childHandle,
                        text, type, true);
                    if (elementItem != null)
                    {
                        elementItem.SetFocus();
                        LogMessage(elementItem.Current.Name + " : " +
                            elementItem.Current.ControlType.ProgrammaticName);
                        if (elementItem.TryGetCurrentPattern(
                            SelectionItemPattern.Pattern, out pattern))
                        {
                            LogMessage("SelectionItemPattern");
                            ((SelectionItemPattern)pattern).AddToSelection();
                        }
                        else
                        {
                            throw new XmlRpcFaultException(123,
                                "Unsupported pattern.");
                        }
                    }
                    else
                    {
                        throw new XmlRpcFaultException(123,
                            "Unable to find the item in list: " + text);
                    }
                }
                return 1;
            }
            catch (Exception ex)
            {
                LogMessage(ex);
                if (ex is XmlRpcFaultException)
                    throw;
                else
                    throw new XmlRpcFaultException(123,
                        "Unhandled exception: " + ex.Message);
            }
            finally
            {
                type = null;
                pattern = null;
                searchTexts = null;
                elementItem = childHandle = null;
            }
        }
        public int MultiRemove(String windowName, String objName,
            String[] texts, bool partialMatch = false)
        {
            if (texts == null)
            {
                throw new XmlRpcFaultException(123, "Argument cannot be empty.");
            }
            Object pattern;
            ControlType[] type;
            string[] searchTexts = texts;
            AutomationElement elementItem;
            List<string> myCollection = new List<string>();
            AutomationElement childHandle = GetObjectHandle(windowName,
                objName);
            if (!utils.IsEnabled(childHandle))
            {
                childHandle = null;
                throw new XmlRpcFaultException(123,
                    "Object state is disabled");
            }
            try
            {
                try
                {
                    childHandle.SetFocus();
                }
                catch (InvalidOperationException ex)
                {
                    LogMessage(ex);
                }
                if (partialMatch)
                {
                    foreach (string text in texts)
                        myCollection.Add("*" + text + "*");
                    // Search for the partial text, rather than the given text
                    searchTexts = myCollection.ToArray();
                }
                type = new ControlType[4] { ControlType.TreeItem,
                    ControlType.ListItem, ControlType.DataItem,
                    ControlType.Custom };
                foreach (string text in searchTexts)
                {
                    elementItem = utils.GetObjectHandle(childHandle,
                        text, type, true);
                    if (elementItem != null)
                    {
                        elementItem.SetFocus();
                        LogMessage(elementItem.Current.Name + " : " +
                            elementItem.Current.ControlType.ProgrammaticName);
                        if (elementItem.TryGetCurrentPattern(
                            SelectionItemPattern.Pattern, out pattern))
                        {
                            LogMessage("SelectionItemPattern");
                            ((SelectionItemPattern)pattern).RemoveFromSelection();
                        }
                        else
                        {
                            throw new XmlRpcFaultException(123,
                                "Unsupported pattern.");
                        }
                    }
                    else
                    {
                        throw new XmlRpcFaultException(123,
                            "Unable to find the item in list: " + text);
                    }
                }
                return 1;
            }
            catch (Exception ex)
            {
                LogMessage(ex);
                if (ex is XmlRpcFaultException)
                    throw;
                else
                    throw new XmlRpcFaultException(123,
                        "Unhandled exception: " + ex.Message);
            }
            finally
            {
                type = null;
                pattern = null;
                searchTexts = null;
                elementItem = childHandle = null;
            }
        }
        public int RightClick(String windowName, String objName, String text)
        {
            if (String.IsNullOrEmpty(text))
            {
                throw new XmlRpcFaultException(123, "Argument cannot be empty.");
            }
            ControlType[] type;
            AutomationElement elementItem;
            AutomationElement childHandle = GetObjectHandle(windowName,
                objName);
            if (!utils.IsEnabled(childHandle))
            {
                childHandle = null;
                throw new XmlRpcFaultException(123,
                    "Object state is disabled");
            }
            Mouse mouse = new Mouse(utils);
            try
            {
                childHandle.SetFocus();
                type = new ControlType[4] { ControlType.TreeItem,
                    ControlType.ListItem, ControlType.DataItem,
                    ControlType.Custom };
                elementItem = utils.GetObjectHandle(childHandle,
                    text, type, true);
                if (elementItem != null)
                {
                    elementItem.SetFocus();
                    LogMessage(elementItem.Current.Name + " : " +
                        elementItem.Current.ControlType.ProgrammaticName);
                    Rect rect = elementItem.Current.BoundingRectangle;
                    mouse.GenerateMouseEvent((int)(rect.X + rect.Width / 2),
                        (int)(rect.Y + rect.Height / 2), "b3c");
                    return 1;
                }
            }
            catch (Exception ex)
            {
                LogMessage(ex);
                if (ex is XmlRpcFaultException)
                    throw;
                else
                    throw new XmlRpcFaultException(123,
                        "Unhandled exception: " + ex.Message);
            }
            finally
            {
                type = null;
                mouse = null;
                elementItem = childHandle = null;
            }
            throw new XmlRpcFaultException(123,
                "Unable to find the item in list: " + text);
        }
        public int ExpandCollapseClick(String windowName, String objName, String text)
        {
            if (String.IsNullOrEmpty(text))
            {
                throw new XmlRpcFaultException(123, "Argument cannot be empty.");
            }
            ControlType[] type;
            AutomationElement elementItem;
            AutomationElement childHandle = GetObjectHandle(windowName,
                objName);
            if (!utils.IsEnabled(childHandle))
            {
                childHandle = null;
                throw new XmlRpcFaultException(123,
                    "Object state is disabled");
            }
            Mouse mouse = new Mouse(utils);
            try
            {
                childHandle.SetFocus();
                type = new ControlType[4] { ControlType.TreeItem,
                    ControlType.ListItem, ControlType.DataItem,
                    ControlType.Custom };
                elementItem = utils.GetObjectHandle(childHandle,
                    text, type, true);
                if (elementItem != null)
                {
                    elementItem.SetFocus();
                    LogMessage(elementItem.Current.Name + " : " +
                        elementItem.Current.ControlType.ProgrammaticName);
                    Rect rect = elementItem.Current.BoundingRectangle;
                    mouse.GenerateMouseEvent((int)(rect.X - 2),
                        (int)(rect.Y + rect.Height / 2), "b1c");
                    return 1;
                }
            }
            catch (Exception ex)
            {
                LogMessage(ex);
                if (ex is XmlRpcFaultException)
                    throw;
                else
                    throw new XmlRpcFaultException(123,
                        "Unhandled exception: " + ex.Message);
            }
            finally
            {
                type = null;
                mouse = null;
                elementItem = childHandle = null;
            }
            throw new XmlRpcFaultException(123,
                "Unable to find the item in list: " + text);
        }
        public int VerifySelectRow(String windowName, String objName,
            String text, bool partialMatch = false)
        {
            if (String.IsNullOrEmpty(text))
            {
                LogMessage("Argument cannot be empty.");
                return 0;
            }
            Object pattern;
            ControlType[] type;
            AutomationElement elementItem;
            AutomationElement childHandle = GetObjectHandle(windowName,
                objName);
            if (!utils.IsEnabled(childHandle))
            {
                childHandle = null;
                LogMessage("Object state is disabled");
                return 0;
            }
            try
            {
                childHandle.SetFocus();
                if (partialMatch)
                    text += "*";
                type = new ControlType[4] { ControlType.TreeItem,
                    ControlType.ListItem, ControlType.DataItem,
                    ControlType.Custom };
                elementItem = utils.GetObjectHandle(childHandle,
                    text, type, true);
                if (elementItem != null)
                {
                    elementItem.SetFocus();
                    LogMessage(elementItem.Current.Name + " : " +
                        elementItem.Current.ControlType.ProgrammaticName);
                    if (elementItem.TryGetCurrentPattern(
                        SelectionItemPattern.Pattern, out pattern))
                    {
                        LogMessage("SelectionItemPattern");
                        if (((SelectionItemPattern)pattern).Current.IsSelected ==
                                true)
                        {
                            LogMessage("Selected");
                            return 1;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                LogMessage(ex);
            }
            finally
            {
                type = null;
                pattern = null;
                elementItem = childHandle = null;
            }
            LogMessage("Unable to find the item in list: " + text);
            return 0;
        }
        public int SelectRowPartialMatch(String windowName, String objName,
            String text)
        {
            return SelectRow(windowName, objName, text, true);
        }
        public int SelectRowIndex(String windowName,
            String objName, int index)
        {
            Object pattern;
            AutomationElement element = null;
            AutomationElement childHandle = GetObjectHandle(windowName,
                objName);
            if (!utils.IsEnabled(childHandle))
            {
                childHandle = null;
                throw new XmlRpcFaultException(123,
                    "Object state is disabled");
            }
            Condition prop1 = new PropertyCondition(
                AutomationElement.ControlTypeProperty, ControlType.ListItem);
            Condition prop2 = new PropertyCondition(
                AutomationElement.ControlTypeProperty, ControlType.TreeItem);
            Condition prop3 = new PropertyCondition(
                AutomationElement.ControlTypeProperty, ControlType.DataItem);
            Condition prop4 = new PropertyCondition(
                AutomationElement.ControlTypeProperty, ControlType.Custom);
            Condition condition = new OrCondition(prop1, prop2, prop3,
                prop4);
            try
            {
                childHandle.SetFocus();
                AutomationElementCollection c = childHandle.FindAll(TreeScope.Children,
                    condition);
                try
                {
                    int columns = GetColumnCount(c);
                    element = c[index * columns];
                    element.SetFocus();
                }
                catch (IndexOutOfRangeException)
                {
                    throw new XmlRpcFaultException(123,
                        "Index out of range: " + index);
                }
                catch (ArgumentException)
                {
                    throw new XmlRpcFaultException(123,
                        "Index out of range: " + index);
                }
                catch (Exception ex)
                {
                    LogMessage(ex);
                    throw new XmlRpcFaultException(123,
                        "Index out of range: " + index);
                }
                if (element != null)
                {
                    LogMessage(element.Current.Name + " : " +
                        element.Current.ControlType.ProgrammaticName);
                    if (element.TryGetCurrentPattern(SelectionItemPattern.Pattern,
                        out pattern))
                    {
                        LogMessage("SelectionItemPattern");
                        element.SetFocus();
                        //((SelectionItemPattern)pattern).Select();
                        // NOTE: Work around, as the above doesn't seem to work
                        // with UIAComWrapper and UIAComWrapper is required
                        // to Edit value in Spin control
                        utils.InternalClick(element);
                        return 1;
                    }
                    else if (element.TryGetCurrentPattern(ExpandCollapsePattern.Pattern,
                        out pattern))
                    {
                        LogMessage("ExpandCollapsePattern");
                        ((ExpandCollapsePattern)pattern).Expand();
                        element.SetFocus();
                        return 1;
                    }
                }
            }
            catch (Exception ex)
            {
                LogMessage(ex);
                if (ex is XmlRpcFaultException)
                    throw;
                else
                    throw new XmlRpcFaultException(123,
                        "Unhandled exception: " + ex.Message);
            }
            finally
            {
                pattern = null;
                element = childHandle = null;
                prop1 = prop2 = prop3 = prop4 = condition = null;
            }
            throw new XmlRpcFaultException(123, "Unable to select item.");
        }
        public int ExpandTableCell(String windowName,
            String objName, int index)
        {
            Object pattern;
            AutomationElement element = null;
            AutomationElement childHandle = GetObjectHandle(windowName,
                objName);
            if (!utils.IsEnabled(childHandle))
            {
                childHandle = null;
                throw new XmlRpcFaultException(123,
                    "Object state is disabled");
            }
            Condition prop1 = new PropertyCondition(
                AutomationElement.ControlTypeProperty, ControlType.ListItem);
            Condition prop2 = new PropertyCondition(
                AutomationElement.ControlTypeProperty, ControlType.TreeItem);
            Condition prop3 = new PropertyCondition(
                AutomationElement.ControlTypeProperty, ControlType.DataItem);
            Condition prop4 = new PropertyCondition(
                AutomationElement.ControlTypeProperty, ControlType.Custom);
            Condition condition = new OrCondition(prop1, prop2, prop3,
                prop4);
            try
            {
                childHandle.SetFocus();
                AutomationElementCollection c = childHandle.FindAll(TreeScope.Children,
                    condition);
                try
                {
                    int columns = GetColumnCount(c);
                    element = c[index * columns];
                }
                catch (IndexOutOfRangeException)
                {
                    throw new XmlRpcFaultException(123, "Index out of range: " + index);
                }
                catch (ArgumentException)
                {
                    throw new XmlRpcFaultException(123, "Index out of range: " + index);
                }
                catch (Exception ex)
                {
                    LogMessage(ex);
                    throw new XmlRpcFaultException(123, "Index out of range: " + index);
                }
                finally
                {
                    c = null;
                    childHandle = null;
                    prop1 = prop2 = prop3 = prop4 = null;
                    condition = null;
                }
                if (element != null)
                {
                    LogMessage(element.Current.Name + " : " +
                        element.Current.ControlType.ProgrammaticName);
                    if (element.TryGetCurrentPattern(ExpandCollapsePattern.Pattern,
                        out pattern))
                    {
                        LogMessage("ExpandCollapsePattern");
                        if (((ExpandCollapsePattern)pattern).Current.ExpandCollapseState ==
                            ExpandCollapseState.Expanded)
                            ((ExpandCollapsePattern)pattern).Collapse();
                        else if (((ExpandCollapsePattern)pattern).Current.ExpandCollapseState ==
                            ExpandCollapseState.Collapsed)
                            ((ExpandCollapsePattern)pattern).Expand();
                        return 1;
                    }
                }
            }
            catch (Exception ex)
            {
                LogMessage(ex);
                if (ex is XmlRpcFaultException)
                    throw;
                else
                    throw new XmlRpcFaultException(123,
                        "Unhandled exception: " + ex.Message);
            }
            finally
            {
                element = null;
                pattern = null;
            }
            throw new XmlRpcFaultException(123, "Unable to expand item.");
        }
        public int SetCellValue(String windowName,
            String objName, int row, int column = 0, String data = null)
        {
            AutomationElement childHandle = GetObjectHandle(windowName,
                objName);
            if (!utils.IsEnabled(childHandle))
            {
                childHandle = null;
                throw new XmlRpcFaultException(123,
                    "Object state is disabled");
            }
            AutomationElement element = null;
            Condition prop1 = new PropertyCondition(
                AutomationElement.ControlTypeProperty, ControlType.ListItem);
            Condition prop2 = new PropertyCondition(
                AutomationElement.ControlTypeProperty, ControlType.TreeItem);
            Condition prop3 = new PropertyCondition(
                AutomationElement.ControlTypeProperty, ControlType.DataItem);
            Condition prop4 = new PropertyCondition(
                AutomationElement.ControlTypeProperty, ControlType.Text);
            Condition prop5 = new PropertyCondition(
                AutomationElement.ControlTypeProperty, ControlType.Custom);
            Condition condition1 = new OrCondition(prop1, prop2, prop3, prop5);
            Condition condition2 = new OrCondition(prop4, prop5);
            try
            {
                childHandle.SetFocus();
                AutomationElementCollection c = childHandle.FindAll(
                    TreeScope.Children, condition1);

                int columns = GetColumnCount(c);
                element = c[row * columns + column];
                c = null;
                if (element != null)
                {
                    if (element.Current.ControlType == ControlType.Text)
                    {
                        throw new XmlRpcFaultException(123,
                            "Not implemented");
                    }
                    else if (element.Current.ControlType == ControlType.TreeItem)
                    {
                        var editorValuePattern = element.GetCurrentPattern(LegacyIAccessiblePattern.Pattern) as LegacyIAccessiblePattern;
                        editorValuePattern.SetValue(data);

                        return 1;
                    }
                    else
                    {
                        // Specific to DataGrid of Windows Forms
                        element.SetFocus();
                        Mouse mouse = new Mouse(utils);
                        Rect rect = element.Current.BoundingRectangle;
                        utils.InternalWait(1);
                        mouse.GenerateMouseEvent((int)(rect.X + rect.Width / 2),
                            (int)(rect.Y + rect.Height / 2), "b1c");
                        utils.InternalWait(1);
                        // Only on second b1c, it becomes edit control
                        // though the edit control is not under current widget
                        // its created in different hierarchy altogether
                        // So, its required to do (from python)
                        // settextvalue("Window Name", "txtEditingControl", "Some text")
                        mouse.GenerateMouseEvent((int)(rect.X + rect.Width / 2),
                            (int)(rect.Y + rect.Height / 2), "b1c");

                        return 1;
                    }
                }
            }
            catch (IndexOutOfRangeException ex)
            {
                LogMessage(ex);
                throw new XmlRpcFaultException(123,
                    "IndexOutOfRangeException: " + "(" + row + ", " + column + "): " + ex);
            }
            catch (ArgumentException ex)
            {
                LogMessage(ex);
                throw new XmlRpcFaultException(123,
                    "ArgumentException: " + "(" + row + ", " + column + "): " + ex);
            }
            catch (Exception ex)
            {
                LogMessage(ex);
                throw new XmlRpcFaultException(123,
                    "Exception: " + "(" + row + ", " + column + "): " + ex);
            }
            finally
            {
                element = childHandle = null;
                prop1 = prop2 = prop3 = prop4 = prop5 = null;
                condition1 = condition2 = null;
            }
            throw new XmlRpcFaultException(123,
                "Unable to set item value.");
        }
        public String GetCellValue(String windowName,
            String objName, int row, int column = 0)
        {
            AutomationElement childHandle = GetObjectHandle(windowName,
                objName);
            if (!utils.IsEnabled(childHandle))
            {
                childHandle = null;
                throw new XmlRpcFaultException(123,
                    "Object state is disabled");
            }
            AutomationElement element = null;
            Condition prop1 = new PropertyCondition(
                AutomationElement.ControlTypeProperty, ControlType.ListItem);
            Condition prop2 = new PropertyCondition(
                AutomationElement.ControlTypeProperty, ControlType.TreeItem);
            Condition prop3 = new PropertyCondition(
                AutomationElement.ControlTypeProperty, ControlType.DataItem);
            Condition prop4 = new PropertyCondition(
                AutomationElement.ControlTypeProperty, ControlType.Text);
            Condition prop5 = new PropertyCondition(
                AutomationElement.ControlTypeProperty, ControlType.Custom);
            Condition condition1 = new OrCondition(prop1, prop2, prop3, prop5);
            Condition condition2 = new OrCondition(prop4, prop5);
            try
            {
                childHandle.SetFocus();
                AutomationElementCollection c = childHandle.FindAll(
                    TreeScope.Children, condition1);
                int columns = GetColumnCount(c);
                element = c[row * columns + column];
                c = null;
                if (element != null)
                {
                    if (element.Current.ControlType == ControlType.Text)
                    {
                        return element.Current.Name;
                    }
                    else if (element.Current.ControlType == ControlType.Custom)
                    {
                        return element.Current.Name;
                    }
                    else if (element.Current.ControlType == ControlType.List)
                    {
                        return element.Current.Name;
                    }
                    else if (element.Current.ControlType == ControlType.ListItem)
                    {
                        return element.Current.Name;
                    }
                    else if (element.Current.ControlType == ControlType.Tree)
                    {
                        return element.Current.Name;
                    }
                    else if (element.Current.ControlType == ControlType.TreeItem)
                    {
                        return element.Current.Name;
                    }
                    else if (element.Current.ControlType == ControlType.Table)
                    {
                        return element.Current.Name;
                    }
                    else if (element.Current.ControlType == ControlType.DataItem)
                    {
                        return element.Current.Name;
                    }
                    else
                    {
                        //try use msaa to access value
                        try
                        {
                            var editorValuePattern = element.GetCurrentPattern(LegacyIAccessiblePattern.Pattern) as LegacyIAccessiblePattern;
                            return editorValuePattern.Current.Value;
                        }
                        catch (Exception ex)
                        {
                            LogMessage(ex);
                        }
                        // Specific to DataGrid of Windows Forms
                        element.SetFocus();
                        Mouse mouse = new Mouse(utils);
                        Rect rect = element.Current.BoundingRectangle;
                        utils.InternalWait(1);
                        mouse.GenerateMouseEvent((int)(rect.X + rect.Width / 2),
                            (int)(rect.Y + rect.Height / 2), "b1c");
                        utils.InternalWait(1);
                        // Only on second b1c, it becomes edit control
                        // though the edit control is not under current widget
                        // its created in different hierarchy altogether
                        // So, its required to do (from python)
                        // gettextvalue("Window Name", "txtEditingControl")
                        mouse.GenerateMouseEvent((int)(rect.X + rect.Width / 2),
                            (int)(rect.Y + rect.Height / 2), "b1c");

                        return "";
                    }
                }
            }
            catch (IndexOutOfRangeException ex)
            {
                LogMessage(ex);
                throw new XmlRpcFaultException(123,
                    "IndexOutOfRangeException: " + "(" + row + ", " + column + "). Ex: " + ex);
            }
            catch (ArgumentException ex)
            {
                LogMessage(ex);
                throw new XmlRpcFaultException(123,
                    "ArgumentException: " + "(" + row + ", " + column + "). Ex: " + ex);
            }
            catch (Exception ex)
            {
                LogMessage(ex);
                throw new XmlRpcFaultException(123,
                    "Exception: " + "(" + row + ", " + column + "). Ex: " + ex);
            }
            finally
            {
                element = childHandle = null;
                prop1 = prop2 = prop3 = prop4 = prop5 = null;
                condition1 = condition2 = null;
            }
            throw new XmlRpcFaultException(123,
                "Unable to get item value.");
        }
        public int[] GetCellSize(String windowName,
            String objName, int row, int column = 0)
        {
            AutomationElement childHandle = GetObjectHandle(windowName,
                objName);
            if (!utils.IsEnabled(childHandle))
            {
                childHandle = null;
                throw new XmlRpcFaultException(123,
                    "Object state is disabled");
            }
            AutomationElement element = null;
            Condition prop1 = new PropertyCondition(
                AutomationElement.ControlTypeProperty, ControlType.ListItem);
            Condition prop2 = new PropertyCondition(
                AutomationElement.ControlTypeProperty, ControlType.TreeItem);
            Condition prop3 = new PropertyCondition(
                AutomationElement.ControlTypeProperty, ControlType.DataItem);
            Condition prop4 = new PropertyCondition(
                AutomationElement.ControlTypeProperty, ControlType.Text);
            Condition prop5 = new PropertyCondition(
                AutomationElement.ControlTypeProperty, ControlType.Custom);
            Condition condition1 = new OrCondition(prop1, prop2, prop3, prop5);
            Condition condition2 = new OrCondition(prop4, prop5);
            try
            {
                childHandle.SetFocus();
                AutomationElementCollection c = childHandle.FindAll(
                    TreeScope.Children, condition1);
                int columns = GetColumnCount(c);
                element = c[row * columns + column];
                c = null;
                if (element != null)
                {
                    Rect rect = childHandle.Current.BoundingRectangle;
                    return new int[] { (int)rect.X, (int)rect.Y,
                        (int)rect.Width, (int)rect.Height };
                }
            }
            catch (IndexOutOfRangeException)
            {
                throw new XmlRpcFaultException(123,
                    "Index out of range: " + "(" + row + ", " + column + ")");
            }
            catch (ArgumentException)
            {
                throw new XmlRpcFaultException(123,
                    "Index out of range: " + "(" + row + ", " + column + ")");
            }
            catch (Exception ex)
            {
                LogMessage(ex);
                throw new XmlRpcFaultException(123,
                    "Index out of range: " + "(" + row + ", " + column + ")");
            }
            finally
            {
                element = childHandle = null;
                prop1 = prop2 = prop3 = prop4 = prop5 = null;
                condition1 = condition2 = null;
            }
            throw new XmlRpcFaultException(123,
                "Unable to get item size.");
        }
        public int GetTableRowIndex(String windowName,
            String objName, String cellValue)
        {
            AutomationElement childHandle = GetObjectHandle(windowName,
                objName);
            if (!utils.IsEnabled(childHandle))
            {
                childHandle = null;
                throw new XmlRpcFaultException(123,
                    "Object state is disabled");
            }
            AutomationElementCollection c1, c2;
            Condition prop1 = new PropertyCondition(
                AutomationElement.ControlTypeProperty, ControlType.ListItem);
            Condition prop2 = new PropertyCondition(
                AutomationElement.ControlTypeProperty, ControlType.TreeItem);
            Condition prop3 = new PropertyCondition(
                AutomationElement.ControlTypeProperty, ControlType.DataItem);
            Condition prop4 = new PropertyCondition(
                AutomationElement.ControlTypeProperty, ControlType.Text);
            Condition prop5 = new PropertyCondition(
                AutomationElement.ControlTypeProperty, ControlType.Custom);
            Condition condition1 = new OrCondition(prop1, prop2, prop3, prop5);
            Condition condition2 = new OrCondition(prop4, prop5);
            try
            {
                childHandle.SetFocus();
                c1 = childHandle.FindAll(TreeScope.Children, condition1);
                int columns = GetColumnCount(c1);
                for (int i = 0; i < c1.Count; i++)
                {
                    if (utils.common.WildcardMatch(c1[i].Current.Name, cellValue))
                        return i / columns;
                }
            }
            catch (Exception ex)
            {
                LogMessage(ex);
            }
            finally
            {
                c1 = c2 = null;
                childHandle = null;
                prop1 = prop2 = prop3 = prop4 = prop5 = null;
                condition1 = condition2 = null;
            }
            throw new XmlRpcFaultException(123,
                    "Unable to get row index: " + cellValue);
        }
        private int GetColumnCount(AutomationElementCollection items)
        {
            // Resolve column count by checking when item's Y coordinate changes from previous item's Y coordinate (no other way, it seems)
            int columns = 0;

            for(int i = 0; i < items.Count; i++)
            {
                if(i > 0 && items[i].Current.BoundingRectangle.Y != items[i-1].Current.BoundingRectangle.Y)
                    break;

                columns++;                        
            }

            return columns;
        }
        public int GetColumnCount(String windowName, String objName)
        {
            AutomationElement childHandle = GetObjectHandle(windowName,
                objName);
            if (!utils.IsEnabled(childHandle))
            {
                childHandle = null;
                throw new XmlRpcFaultException(123,
                    "Object state is disabled");
            }
            AutomationElementCollection c;
            Condition prop1 = new PropertyCondition(
                AutomationElement.ControlTypeProperty, ControlType.ListItem);
            Condition prop2 = new PropertyCondition(
                AutomationElement.ControlTypeProperty, ControlType.TreeItem);
            Condition prop3 = new PropertyCondition(
                AutomationElement.ControlTypeProperty, ControlType.DataItem);
            Condition prop4 = new PropertyCondition(
                AutomationElement.ControlTypeProperty, ControlType.Custom);
            Condition condition = new OrCondition(prop1, prop2, prop3, prop4);
            try
            {
                c = childHandle.FindAll(TreeScope.Children, condition);
                if (c == null)
                    throw new XmlRpcFaultException(123,
                        "Unable to get row count.");

                int columns = GetColumnCount(c);                    
                return columns;
            }
            catch (Exception ex)
            {
                LogMessage(ex);
                if (ex is XmlRpcFaultException)
                    throw;
                else
                    throw new XmlRpcFaultException(123,
                        "Unhandled exception: " + ex.Message);
            }
            finally
            {
                c = null;
                childHandle = null;
                prop1 = prop2 = prop3 = prop4 = condition = null;
            }
        }
        public int GetRowCount(String windowName, String objName)
        {
            AutomationElement childHandle = GetObjectHandle(windowName,
                objName);
            if (!utils.IsEnabled(childHandle))
            {
                childHandle = null;
                throw new XmlRpcFaultException(123,
                    "Object state is disabled");
            }
            AutomationElementCollection c;
            Condition prop1 = new PropertyCondition(
                AutomationElement.ControlTypeProperty, ControlType.ListItem);
            Condition prop2 = new PropertyCondition(
                AutomationElement.ControlTypeProperty, ControlType.TreeItem);
            Condition prop3 = new PropertyCondition(
                AutomationElement.ControlTypeProperty, ControlType.DataItem);
            Condition prop4 = new PropertyCondition(
                AutomationElement.ControlTypeProperty, ControlType.Custom);
            Condition condition = new OrCondition(prop1, prop2, prop3, prop4);
            try
            {
                c = childHandle.FindAll(TreeScope.Children, condition);
                if (c == null)
                    throw new XmlRpcFaultException(123,
                        "Unable to get row count.");

                int columns = GetColumnCount(c);

                if(columns == 0)
                    return 0;

                return c.Count / columns;
            }
            catch (Exception ex)
            {
                LogMessage(ex);
                if (ex is XmlRpcFaultException)
                    throw;
                else
                    throw new XmlRpcFaultException(123,
                        "Unhandled exception: " + ex.Message);
            }
            finally
            {
                c = null;
                childHandle = null;
                prop1 = prop2 = prop3 = prop4 = condition = null;
            }
        }
        public int SingleClickRow(String windowName, String objName, String text)
        {
            return ClickRow(windowName, objName, text, "b1c");
        }
        int ClickRow(String windowName, String objName, String text, string clickType)
        {
            if (String.IsNullOrEmpty(text))
            {
                throw new XmlRpcFaultException(123, "Argument cannot be empty.");
            }
            Object pattern;
            ControlType[] type;
            AutomationElement elementItem;
            AutomationElement childHandle = GetObjectHandle(windowName,
                objName);
            if (!utils.IsEnabled(childHandle))
            {
                childHandle = null;
                throw new XmlRpcFaultException(123,
                    "Object state is disabled");
            }
            try
            {
                try
                {
                    childHandle.SetFocus();
                }
                catch (InvalidOperationException ex)
                {
                    LogMessage(ex);
                }
                type = new ControlType[4] { ControlType.TreeItem,
                    ControlType.ListItem, ControlType.DataItem,
                    ControlType.Custom };
                elementItem = utils.GetObjectHandle(childHandle,
                    text, type, true);
                if (elementItem != null)
                {
                    elementItem.SetFocus();
                    LogMessage(elementItem.Current.Name + " : " +
                        elementItem.Current.ControlType.ProgrammaticName);
                    if (elementItem.TryGetCurrentPattern(
                        SelectionItemPattern.Pattern, out pattern))
                    {
                        LogMessage("SelectionItemPattern");
                        //((SelectionItemPattern)pattern).Select();
                        // NOTE: Work around, as the above doesn't seem to work
                        // with UIAComWrapper and UIAComWrapper is required
                        // to Edit value in Spin control
                        Mouse mouse = new Mouse(utils);
                        Rect rect = elementItem.Current.BoundingRectangle;
                        mouse.GenerateMouseEvent((int)(rect.X + rect.Width / 2),
                            (int)(rect.Y + rect.Height / 2), clickType);
                        return 1;
                    }
                    else if (elementItem.TryGetCurrentPattern(
                        ExpandCollapsePattern.Pattern, out pattern))
                    {
                        LogMessage("ExpandCollapsePattern");
                        ((ExpandCollapsePattern)pattern).Expand();
                        return 1;
                    }
                    else
                    {
                        throw new XmlRpcFaultException(123,
                            "Unsupported pattern.");
                    }
                }
            }
            catch (Exception ex)
            {
                LogMessage(ex);
                if (ex is XmlRpcFaultException)
                    throw;
                else
                    throw new XmlRpcFaultException(123,
                        "Unhandled exception: " + ex.Message);
            }
            finally
            {
                type = null;
                pattern = null;
                elementItem = childHandle = null;
            }
            throw new XmlRpcFaultException(123,
                "Unable to find the item in list: " + text);
        }
        public int DoubleClickRow(String windowName, String objName, String text)
        {
            return ClickRow(windowName, objName, text, "b1d");
        }

        public int DoubleClickRowIndex(String windowName, String objName,
            int row, int column = 0)
        {
            AutomationElement childHandle = GetObjectHandle(windowName,
                objName);
            if (!utils.IsEnabled(childHandle))
            {
                childHandle = null;
                throw new XmlRpcFaultException(123,
                    "Object state is disabled");
            }
            AutomationElement element = null;
            Condition prop1 = new PropertyCondition(
                AutomationElement.ControlTypeProperty, ControlType.ListItem);
            Condition prop2 = new PropertyCondition(
                AutomationElement.ControlTypeProperty, ControlType.TreeItem);
            Condition prop3 = new PropertyCondition(
                AutomationElement.ControlTypeProperty, ControlType.DataItem);
            Condition prop4 = new PropertyCondition(
                AutomationElement.ControlTypeProperty, ControlType.Text);
            Condition prop5 = new PropertyCondition(
                AutomationElement.ControlTypeProperty, ControlType.Custom);
            Condition condition1 = new OrCondition(prop1, prop2, prop3, prop5);
            Condition condition2 = new OrCondition(prop4, prop5);
            try
            {
                childHandle.SetFocus();
                AutomationElementCollection c = childHandle.FindAll(
                    TreeScope.Children, condition1);
                int columns = GetColumnCount(c);
                element = c[row * columns + column];
                c = null;
                if (element != null)
                {
                    Mouse mouse = new Mouse(utils);
                    Rect rect = element.Current.BoundingRectangle;
                    mouse.GenerateMouseEvent((int)(rect.X + rect.Width / 2),
                        (int)(rect.Y + rect.Height / 2), "b1d");
                    return 1;
                }
            }
            catch (IndexOutOfRangeException)
            {
                throw new XmlRpcFaultException(123,
                    "Index out of range: " + "(" + row + ", " + column + ")");
            }
            catch (ArgumentException)
            {
                throw new XmlRpcFaultException(123,
                    "Index out of range: " + "(" + row + ", " + column + ")");
            }
            catch (Exception ex)
            {
                LogMessage(ex);
                throw new XmlRpcFaultException(123,
                    "Index out of range: " + "(" + row + ", " + column + ")");
            }
            finally
            {
                element = childHandle = null;
                prop1 = prop2 = prop3 = prop4 = prop5 = null;
                condition1 = condition2 = null;
            }
            throw new XmlRpcFaultException(123,
                "Unable to find the item in list: " + row);
        }
    }
}
