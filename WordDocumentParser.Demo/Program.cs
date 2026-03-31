using System;
using System.IO;
using System.Linq;
using WordDocumentParser.Demo.Features.Concatenation;
using WordDocumentParser.Demo.Features.ContentControls;
using WordDocumentParser.Demo.Features.DocumentCreation;
using WordDocumentParser.Demo.Features.DocumentProperties;
using WordDocumentParser.Demo.Features.Examples;
using WordDocumentParser.Demo.Features.Parsing;
using WordDocumentParser.Demo.Features.RoundTrip;
using WordDocumentParser.Demo.Features.Styles;
using WordDocumentParser.Demo.Features.Tables;
using WordDocumentParser.Demo.Features.Fonts;

namespace WordDocumentParser.Demo;

/// <summary>
/// Demonstration program showing how to use the Word Document Tree Parser and Writer library.
/// </summary>
class Program
{
    static void Main(string[] args)
    {
        string inputDoc = @"C:\isolated\tb_mod_test.docx";

        TableModificationDemo.Run(inputDoc);
    }
}
