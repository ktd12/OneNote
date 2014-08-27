using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Navigation;
using Microsoft.Phone.Controls;
using Microsoft.Phone.Shell;
using System.Windows.Data;
using System.Text.RegularExpressions;

namespace CreateOneNotePage
{
    /// <summary>
    /// This example demonstrates how to create a page in a specific notebook and section
    /// If either the notebook or section that the user specifies does not exist, it will be created
    /// </summary>
    public partial class MainPage : PhoneApplicationPage
    {
        private clsOneNote _myOneNote;
        /// <summary>
        /// This instance contains all of the one note routines and data structures
        /// It is also used as the datacontext for this page
        /// </summary>
        public clsOneNote myOneNote
        {
            get { return _myOneNote; }
            set { _myOneNote = value; }
        }
        // Constructor
        public MainPage()
        {
            InitializeComponent();
            myOneNote = new clsOneNote();
            this.DataContext = myOneNote;         
        }
        /// <summary>
        /// Create a page in the Notebook and Section specified
        /// If either the Notebook or Section do not exist, they are created
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnCreatePage_Tap(object sender, System.Windows.Input.GestureEventArgs e)
        {
            if (myOneNote.NotebookName == String.Empty || myOneNote.SectionName == String.Empty 
                || myOneNote.NotebookName == clsOneNote.cNotebookNameHint || myOneNote.SectionName == clsOneNote.cSectionNameHint)
            {
                MessageBox.Show("Notebook name and Section name must be entered", "Error", MessageBoxButton.OK);
                return;
            }

        
            string simpleHtml = "<html>" +
                        "<head>" +
                        "<title> A page created from basic HTML-formatted text </title>" +
                        "<meta name=\"created\" content=\"" + DateTime.Now.ToShortDateString()+ "\" />" +
                        "</head>" +
                        "<body>" +
                        "<p>This is a page that just contains some simple <i>formatted</i> <b>text</b></p>" +
                        "<p>" + DateTime.Now.ToString() + "</p>" +
                        "</body>" +
                        "</html>";

           myOneNote.CreateSimpleHtml(simpleHtml);

        }
        /// <summary>
        /// Display the page that was just created using the OneNote Client
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void hyperlinkViewNote_Tap(object sender, System.Windows.Input.GestureEventArgs e)
        {
            myOneNote.HyperlinkButton_Click();
        }

      
        private void txtNotebookName_GotFocus(object sender, RoutedEventArgs e)
        {
            string currentFieldValue = txtNotebookName.Text;
            if (clsOneNote.cNotebookNameHint.Equals(currentFieldValue))
            {
                txtNotebookName.Text = string.Empty;
            }
        }

        private void txtNotebookName_LostFocus(object sender, RoutedEventArgs e)
        {
            TextBox textBox = sender as TextBox;
            string newValue = textBox.Text.Trim();

            if (newValue == clsOneNote.cNotebookNameHint)
            { newValue = String.Empty; }

            if (HasIllegalCharacters(newValue))
            {
                MessageBox.Show("Only number, letters and spaces are allowed", "Illegal Name", MessageBoxButton.OK);
                newValue = String.Empty; 
            }

            if (String.IsNullOrWhiteSpace(newValue))
            {
                textBox.Text = clsOneNote.cNotebookNameHint;
                BindingExpression bindingExpr = textBox.GetBindingExpression(TextBox.TextProperty);
                bindingExpr.UpdateSource();
              
            }
            else
            {
                textBox.Text = newValue;
                BindingExpression bindingExpr = textBox.GetBindingExpression(TextBox.TextProperty);
                bindingExpr.UpdateSource();
            }
        }

      
        private void txtSectionName_GotFocus(object sender, RoutedEventArgs e)
        {
            string currentFieldValue = txtSectionName.Text;
            if (clsOneNote.cSectionNameHint.Equals(currentFieldValue))
            {
                txtSectionName.Text = string.Empty;
            }
        }

        private void txtSectionName_LostFocus(object sender, RoutedEventArgs e)
        {
            TextBox textBox = sender as TextBox;
            string newValue = textBox.Text.Trim();

            if (HasIllegalCharacters(newValue))
            {
                MessageBox.Show("Only numbers, letters and spaces are allowed", "Illegal Name", MessageBoxButton.OK);
                newValue = String.Empty; 
            }

            if (newValue == clsOneNote.cSectionNameHint)
            { newValue = String.Empty; }

            if (String.IsNullOrWhiteSpace(newValue))
            {
                textBox.Text = clsOneNote.cSectionNameHint;
                BindingExpression bindingExpr = textBox.GetBindingExpression(TextBox.TextProperty);
                bindingExpr.UpdateSource();
            }
            else
            {
                textBox.Text = newValue;
                BindingExpression bindingExpr = textBox.GetBindingExpression(TextBox.TextProperty);
                bindingExpr.UpdateSource();
            }
        }

    
        private void OnSessionChanged(object sender, Microsoft.Live.Controls.LiveConnectSessionChangedEventArgs e)
        {
            myOneNote.OnSessionChanged(sender, e);
        }
        
        private Boolean HasIllegalCharacters(string strIn)
        {
            return Regex.IsMatch(strIn, @"[^\w\ ]");
        }
    }
}