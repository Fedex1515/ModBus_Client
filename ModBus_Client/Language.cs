using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using System.Windows.Controls;
using System.IO;
using System.Web.Script.Serialization;

namespace LanguageLib
{
    class Language
    {
        public dynamic languageTemplate;
        System.Windows.Window win;

        public Language(System.Windows.Window win_)
        {
            win = win_;
        }

        public void loadLanguageTemplate(string templateName)
        {
            string file_content = File.ReadAllText("Lang/" + templateName + ".json");

            JavaScriptSerializer jss = new JavaScriptSerializer();
            languageTemplate = jss.Deserialize<dynamic>(file_content);

            foreach (KeyValuePair<string, dynamic> group in languageTemplate)
            {
                // debug
                // Console.WriteLine("Group: " + group.Key + ": " + group.Value);

                switch (group.Key)
                {
                    // LABELS
                    case "labels":
                        // debug
                        // Console.WriteLine("Fould label group");

                        foreach (KeyValuePair<string, dynamic> item in group.Value)
                        {
                            // debug
                            // Console.WriteLine("Value: " + item.Key + ": " + item.Value);

                            var myLabel = (Label)win.FindName(item.Key);

                            if (myLabel != null)
                            {
                                if (item.Value is string)
                                {
                                    myLabel.Content = item.Value;
                                }
                                else
                                {
                                    foreach (KeyValuePair<string, dynamic> prop in item.Value)
                                    {
                                        switch (prop.Key)
                                        {
                                            case "label":
                                                myLabel.Content = prop.Value;
                                                break;

                                            case "toolTip":
                                                myLabel.ToolTip = prop.Value;
                                                break;
                                        }
                                    }
                                }
                            }
                        }
                        break;

                    // RADIOBUTTONS
                    case "radioButtons":

                        foreach (KeyValuePair<string, dynamic> item in group.Value)
                        {
                            var myRadioButton = (RadioButton)win.FindName(item.Key);

                            if (myRadioButton != null)
                            {
                                if (item.Value is string)
                                {
                                    myRadioButton.Content = item.Value;
                                }
                                else
                                {
                                    foreach (KeyValuePair<string, dynamic> prop in item.Value)
                                    {
                                        switch (prop.Key)
                                        {
                                            case "label":
                                                myRadioButton.Content = prop.Value;
                                                break;

                                            case "toolTip":
                                                myRadioButton.ToolTip = prop.Value;
                                                break;
                                        }
                                    }
                                }
                            }
                        }
                        break;

                    // CHECKBOXES
                    case "checkBoxes":

                        foreach (KeyValuePair<string, dynamic> item in group.Value)
                        {
                            var myCheckBox = (CheckBox)win.FindName(item.Key);

                            if (myCheckBox != null)
                            {
                                if (item.Value is string)
                                {
                                    myCheckBox.Content = item.Value;
                                }
                                else
                                {
                                    foreach (KeyValuePair<string, dynamic> prop in item.Value)
                                    {
                                        switch (prop.Key)
                                        {
                                            case "label":
                                                myCheckBox.Content = prop.Value;
                                                break;

                                            case "toolTip":
                                                myCheckBox.ToolTip = prop.Value;
                                                break;
                                        }
                                    }
                                }
                            }
                        }
                        break;

                    // BUTTONS
                    case "buttons":

                        foreach (KeyValuePair<string, dynamic> item in group.Value)
                        {
                            var myButton = (Button)win.FindName(item.Key);

                            if (myButton != null)
                            {
                                if (item.Value is string)
                                {
                                    myButton.Content = item.Value;
                                }
                                else
                                {
                                    foreach (KeyValuePair<string, dynamic> prop in item.Value)
                                    {
                                        switch (prop.Key)
                                        {
                                            case "label":
                                                myButton.Content = prop.Value;
                                                break;

                                            case "toolTip":
                                                myButton.ToolTip = prop.Value;
                                                break;
                                        }
                                    }
                                }
                            }
                        }
                        break;

                    // TEXTBLOXK
                    case "textBlocks":

                        foreach (KeyValuePair<string, dynamic> item in group.Value)
                        {
                            var myTextBlock = (TextBlock)win.FindName(item.Key);

                            if (myTextBlock != null)
                            {
                                if (item.Value is string)
                                {
                                    myTextBlock.Text = item.Value;
                                }
                                else
                                {
                                    foreach (KeyValuePair<string, dynamic> prop in item.Value)
                                    {
                                        switch (prop.Key)
                                        {
                                            case "label":
                                                myTextBlock.Text = prop.Value;
                                                break;

                                            case "toolTip":
                                                myTextBlock.ToolTip = prop.Value;
                                                break;
                                        }
                                    }
                                }
                            }
                        }
                        break;

                    // TEXTBOX
                    case "textBoxes":

                        foreach (KeyValuePair<string, dynamic> item in group.Value)
                        {
                            var myTextBox = (TextBox)win.FindName(item.Key);

                            if (myTextBox != null)
                            {
                                if (item.Value is string)
                                {
                                    myTextBox.Text = item.Value;
                                }
                                else
                                {
                                    foreach (KeyValuePair<string, dynamic> prop in item.Value)
                                    {
                                        switch (prop.Key)
                                        {
                                            case "label":
                                                myTextBox.Text = prop.Value;
                                                break;

                                            case "toolTip":
                                                myTextBox.ToolTip = prop.Value;
                                                break;
                                        }
                                    }
                                }
                            }
                        }
                        break;

                    // PICTUREBOX:
                    case "pictureBoxes":
                    case "borders":

                        foreach (KeyValuePair<string, dynamic> item in group.Value)
                        {
                            var myPictureBox = (Border)win.FindName(item.Key);

                            if (myPictureBox != null)
                            {
                                if (item.Value is string)
                                {
                                    ;// myPictureBox.Text = item.Value;
                                }
                                else
                                {
                                    foreach (KeyValuePair<string, dynamic> prop in item.Value)
                                    {
                                        switch (prop.Key)
                                        {
                                            case "label":
                                                ;// myPictureBox.Text = prop.Value;
                                                break;

                                            case "toolTip":
                                                myPictureBox.ToolTip = prop.Value;
                                                break;
                                        }
                                    }
                                }
                            }
                        }
                        break;

                    // COMBOBOXES:
                    case "comboBoxes":

                        foreach (KeyValuePair<string, dynamic> item in group.Value)
                        {
                            var myPictureBox = (ComboBox)win.FindName(item.Key);

                            if (myPictureBox != null)
                            {
                                if (item.Value is string)
                                {
                                    ;// myPictureBox.Text = item.Value;
                                }
                                else
                                {
                                    foreach (KeyValuePair<string, dynamic> prop in item.Value)
                                    {
                                        switch (prop.Key)
                                        {
                                            case "label":
                                                ;// myPictureBox.Text = prop.Value;
                                                break;

                                            case "toolTip":
                                                myPictureBox.ToolTip = prop.Value;
                                                break;
                                        }
                                    }
                                }
                            }
                        }
                        break;

                    // TABITEMS
                    case "tabItems":

                        foreach (KeyValuePair<string, dynamic> item in group.Value)
                        {
                            var myTab = (TabItem)win.FindName(item.Key);

                            if (myTab != null)
                            {
                                if (item.Value is string)
                                {
                                    myTab.Header = item.Value;
                                }
                                else
                                {
                                    foreach (KeyValuePair<string, dynamic> prop in item.Value)
                                    {
                                        switch (prop.Key)
                                        {
                                            case "label":
                                                myTab.Header = prop.Value;
                                                break;

                                            case "toolTip":
                                                myTab.ToolTip = prop.Value;
                                                break;
                                        }
                                    }
                                }
                            }
                        }
                        break;

                    // MENUITEMS
                    case "menuItems":

                        foreach (KeyValuePair<string, dynamic> item in group.Value)
                        {
                            var myMenu = (MenuItem)win.FindName(item.Key);

                            if (myMenu != null)
                            {
                                if (item.Value is string)
                                {
                                    myMenu.Header = item.Value;
                                }
                                else
                                {
                                    foreach (KeyValuePair<string, dynamic> prop in item.Value)
                                    {
                                        switch (prop.Key)
                                        {
                                            case "label":
                                                myMenu.Header = prop.Value;
                                                break;

                                            case "toolTip":
                                                myMenu.ToolTip = prop.Value;
                                                break;
                                        }
                                    }
                                }
                            }
                        }
                        break;

                    // STRING
                    case "strings":
                        ;
                        break;
                }
            }
        }
    }
}
