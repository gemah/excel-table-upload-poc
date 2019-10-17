using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using TableDataAPI.Models;
using TableDataAPI.Services;

namespace TableDataAPI.Controllers
{
    /// <summary>
    /// The controler for the table reading and parsing logic.
    /// </summary>
    [Route("api/[controller]")]
    [ApiController]
    public class TableController : ControllerBase
    {
        /// <summary>
        /// Empty constructor.
        /// </summary>
        public TableController() { }

        /// <summary>
        /// Basic GET endpoint for checking if the server is currently running.
        /// </summary>
        /// <returns></returns>
        [HttpGet]
        public ActionResult<string> Get()
        {
            return "Table Parsing service is up and running.";
        }

        /// <summary>
        /// POST that shall receive MS Excel Table files (XLS(x)) and will parse the results
        /// into readable content (json)
        /// </summary>
        /// <param name="data">The table data input.</param>
        /// <returns></returns>
        [HttpPost]
        [Consumes("multipart/form-data", "application/json")]
        [Produces("application/json")]
        public ActionResult<IEnumerable<Content>> Post([FromForm] object data)
        {
            List<Content> contents = new List<Content>();
            string fileFormat = null;
            IFormFile file = null;

            if (Request.Form.Files.Count > 0)
            {
                file = Request.Form.Files.First();
                Console.WriteLine("FILE INFO:\nContent-Type: " + file.ContentType + "\nFilename: " + file.FileName + "\nLength: " + file.Length);
                fileFormat = file.FileName.Split(".").Last();
            }

            // TEST - JSON Content Parsing (Will only work if reading data from Body)
            //if (data is JObject)
            //{
            //    dynamic obj = JsonConvert.DeserializeObject(data.ToString());
            //    contents.Add(new Content()
            //    {
            //        Name = obj.Name,
            //        Value = obj.Value,
            //        Comment = obj.Comment
            //    });
            //}

            if (file.Length > 0)
            {
                using (MemoryStream stream = new MemoryStream())
                {
                    ISheet excelSheet;
                    file.CopyTo(stream);
                    stream.Position = 0;

                    if (fileFormat.Equals("xls"))
                    {
                        excelSheet = new HSSFWorkbook(stream).GetSheetAt(0);
                    }
                    else if (fileFormat.Equals("xlsx"))
                    {
                        excelSheet = new XSSFWorkbook(stream).GetSheetAt(0);
                    }
                    else
                    {
                        throw new Exception("INVALID FILE FORMAT!");
                    }

                    IRow header = excelSheet.GetRow(0);
                    string[] headerNames = new string[header.Cells.Count];

                    for (int i = 0; i < headerNames.Length; i++)
                    {
                        headerNames[i] = header.GetCell(i).StringCellValue;
                    }

                    List<string[]> values = new List<string[]>();
                    string[] line;

                    for (int i = 1; i < (excelSheet.LastRowNum + 1); i++)
                    {
                        IList<ICell> currentLine = excelSheet.GetRow(i).Cells;
                        line = new string[currentLine.Count];

                        for (int j = 0; j < line.Length; j++)
                        {
                            line[j] = (currentLine[j].CellType == CellType.String) ? currentLine[j].StringCellValue : currentLine[j].NumericCellValue.ToString();
                        }

                        values.Add(line);
                    }

                    ContentTranslator translator = new ContentTranslator(headerNames);
                    contents = translator.Translate(values);
                }
            }

            return contents.Count > 0 ? StatusCode((int)HttpStatusCode.OK, contents) : StatusCode((int)HttpStatusCode.BadRequest, null);
        }

    }
}