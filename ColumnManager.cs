using System;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Outlook;
using SysException = System.Exception;

namespace Outlook_Purview_Sensitivity
{
    internal static class ColumnManager
    {
        public const string FieldName = "PurviewLabel";

        public static void EnsureColumn(MAPIFolder folder)
        {
            if (folder == null) return;

            try
            {
                UserDefinedProperties props = folder.UserDefinedProperties;
                bool found = false;
                for (int i = 1; i <= props.Count; i++)
                {
                    UserDefinedProperty prop = props[i];
                    if (prop.Name == FieldName)
                    {
                        found = true;
                        Marshal.ReleaseComObject(prop);
                        break;
                    }
                    Marshal.ReleaseComObject(prop);
                }

                if (!found)
                    props.Add(FieldName, OlUserPropertyType.olText);

                Marshal.ReleaseComObject(props);

                object view = folder.CurrentView;
                if (view is TableView tableView)
                {
                    ViewFields fields = tableView.ViewFields;
                    bool columnFound = false;
                    for (int i = 1; i <= fields.Count; i++)
                    {
                        ViewField field = fields[i];
                        if (field.ViewXMLSchemaName == FieldName)
                        {
                            columnFound = true;
                            Marshal.ReleaseComObject(field);
                            break;
                        }
                        Marshal.ReleaseComObject(field);
                    }

                    if (!columnFound)
                    {
                        fields.Add(FieldName);
                        tableView.Save();
                    }

                    Marshal.ReleaseComObject(fields);
                    Marshal.ReleaseComObject(tableView);
                }

                Marshal.ReleaseComObject(view);
            }
            catch (SysException)
            {
            }
        }

        public static void StampItem(MailItem mailItem)
        {
            if (mailItem == null) return;

            try
            {
                string labelName = LabelReader.GetLabelName(mailItem);

                try
                {
                    UserProperty existing = mailItem.UserProperties[FieldName];
                    if (existing != null)
                    {
                        string val = existing.Value as string;
                        Marshal.ReleaseComObject(existing);
                        if (val == labelName) return;
                    }
                }
                catch
                {
                }

                UserProperty userProp = mailItem.UserProperties.Add(
                    FieldName, OlUserPropertyType.olText);
                userProp.Value = labelName;
                mailItem.Save();
                Marshal.ReleaseComObject(userProp);
            }
            catch (SysException)
            {
            }
        }

        public static void StampFolder(MAPIFolder folder, int maxItems = 50)
        {
            if (folder == null) return;

            try
            {
                Items items = folder.Items;
                int count = 0;
                for (int i = 1; i <= items.Count && count < maxItems; i++)
                {
                    object item = items[i];
                    if (item is MailItem mailItem)
                    {
                        string label = LabelReader.GetLabelName(mailItem);
                        if (label != "None")
                        {
                            StampItem(mailItem);
                            count++;
                        }
                        Marshal.ReleaseComObject(mailItem);
                    }
                    Marshal.ReleaseComObject(item);
                }
                Marshal.ReleaseComObject(items);

                // Refresh view to show new values
                object view = folder.CurrentView;
                if (view is TableView tableView)
                {
                    tableView.Apply();
                    Marshal.ReleaseComObject(tableView);
                }
                Marshal.ReleaseComObject(view);
            }
            catch (SysException)
            {
            }
        }
    }
}
