using AngleSharp.Browser;
using AngleSharp.Dom;
using AngleSharp.Html.Parser;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Elib
{
    public partial class FormAuthors : Form
    {
        IEnumerable<IElement> _authors=null;
        public FormAuthors(IEnumerable<IElement>authors)
        {
            InitializeComponent();
            foreach(var author in authors)
            {
                listBox1.Items.Add(author.Children[2].Children[0].TextContent);
            }
            _authors = authors;
        }
        public Educator SelectedEducator;
        private void button1_Click(object sender, EventArgs e)
        {
            if(listBox1.SelectedItem!=null)
            {
                var a = _authors.Where(x => x.TextContent.Contains(listBox1.SelectedItem.ToString()));

                HtmlParser Parser = new HtmlParser();
                var doc = Parser.ParseDocument(a.First().OuterHtml);

                SelectedEducator = new Educator();
                SelectedEducator.Id = Regex.Match(doc.Body.Children[2].Children[0].Id, @"(\d+)").Groups[1].Value;
                SelectedEducator.FIO = doc.Body.Children[3].TextContent.Trim();
                SelectedEducator.LinkOnPublish = $"https://www.elibrary.ru/author_items.asp?authorid={SelectedEducator.Id}&pubrole=100&show_refs=1&show_option=0";

                DialogResult = DialogResult.OK;
            }
            else
            {
                MessageBox.Show("Не выбран преподаватель","Ошибка");
            }
        }
    }
}
