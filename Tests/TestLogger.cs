using System;
using ExcelOrderAddIn.Logging;
using NUnit.Framework;
using Tests.Stubs;

namespace Tests
{
    public class TestLogger
    {
        [Test]
        public void GetLogAsRichTextWithoutEntryNumbers()
        {
            var dateTimeForTests = new DateTimeForTests();
            var logger = new Logger(5, dateTimeForTests);

            dateTimeForTests.Now = new DateTime(2020, 6, 30, 7, 25, 30);
            logger.Info("first");
            dateTimeForTests.Now = new DateTime(2020, 6, 30, 7, 25, 31);
            logger.Info("second");
            dateTimeForTests.Now = new DateTime(2020, 6, 30, 7, 25, 32);
            logger.Info("third");

            var logAsRichText = logger.GetLogAsRichText(false);

            Assert.AreEqual(
                "{\\rtf1{\\colortbl;\\red0\\green0\\blue0;}\r\n" +
                "\\cf1 6/30/2020 7:25:30 AM: \\cf1 first\\par\r\n" +
                "\\cf1 6/30/2020 7:25:31 AM: \\cf1 second\\par\r\n" +
                "\\cf1 6/30/2020 7:25:32 AM: \\cf1 third\\par\r\n",
                logAsRichText);
        }

        [Test]
        public void GetLogAsRichTextWithEntryNumbers()
        {
            var dateTimeForTests = new DateTimeForTests();
            var logger = new Logger(5, dateTimeForTests);

            dateTimeForTests.Now = new DateTime(2020, 6, 30, 7, 25, 30);
            logger.Info("first");
            dateTimeForTests.Now = new DateTime(2020, 6, 30, 7, 25, 31);
            logger.Info("second");
            dateTimeForTests.Now = new DateTime(2020, 6, 30, 7, 25, 32);
            logger.Info("third");

            var logAsRichText = logger.GetLogAsRichText(true);

            Assert.AreEqual(
                "{\\rtf1{\\colortbl;\\red0\\green0\\blue0;}\r\n" +
                "\\cf1 1. \\cf1 6/30/2020 7:25:30 AM: \\cf1 first\\par\r\n" +
                "\\cf1 2. \\cf1 6/30/2020 7:25:31 AM: \\cf1 second\\par\r\n" +
                "\\cf1 3. \\cf1 6/30/2020 7:25:32 AM: \\cf1 third\\par\r\n",
                logAsRichText);
        }

        [Test]
        public void GetLogAsRichTextWithOverflow()
        {
            var dateTimeForTests = new DateTimeForTests();
            var logger = new Logger(3, dateTimeForTests);

            dateTimeForTests.Now = new DateTime(2020, 6, 30, 7, 25, 30);
            logger.Info("first");
            dateTimeForTests.Now = new DateTime(2020, 6, 30, 7, 25, 31);
            logger.Info("second");
            dateTimeForTests.Now = new DateTime(2020, 6, 30, 7, 25, 32);
            logger.Info("third");
            dateTimeForTests.Now = new DateTime(2020, 6, 30, 7, 25, 33);
            logger.Info("fourth");
            dateTimeForTests.Now = new DateTime(2020, 6, 30, 7, 25, 34);
            logger.Info("fifth");

            var logAsRichText = logger.GetLogAsRichText(false);

            Assert.AreEqual(
                "{\\rtf1{\\colortbl;\\red0\\green0\\blue0;}\r\n" +
                "\\cf1 6/30/2020 7:25:32 AM: \\cf1 third\\par\r\n" +
                "\\cf1 6/30/2020 7:25:33 AM: \\cf1 fourth\\par\r\n" +
                "\\cf1 6/30/2020 7:25:34 AM: \\cf1 fifth\\par\r\n",
                logAsRichText);
        }

        [Test]
        public void GetLogAsRichTextWithError()
        {
            var dateTimeForTests = new DateTimeForTests();
            var logger = new Logger(5, dateTimeForTests);

            dateTimeForTests.Now = new DateTime(2020, 6, 30, 7, 25, 30);
            logger.Info("first");
            dateTimeForTests.Now = new DateTime(2020, 6, 30, 7, 25, 31);
            logger.Error("second");
            dateTimeForTests.Now = new DateTime(2020, 6, 30, 7, 25, 32);
            logger.Info("third");

            var logAsRichText = logger.GetLogAsRichText(false);

            Assert.AreEqual(
                "{\\rtf1{\\colortbl;\\red0\\green0\\blue0;\\red255\\green0\\blue0;}\r\n" +
                "\\cf1 6/30/2020 7:25:30 AM: \\cf1 first\\par\r\n" +
                "\\cf1 6/30/2020 7:25:31 AM: \\cf2 second\\par\r\n" +
                "\\cf1 6/30/2020 7:25:32 AM: \\cf1 third\\par\r\n",
                logAsRichText);
        }

        [Test]
        public void ClearClearsLogsAndUnnecessaryColors()
        {
            var dateTimeForTests = new DateTimeForTests();
            var logger = new Logger(5, dateTimeForTests);

            dateTimeForTests.Now = new DateTime(2020, 6, 30, 7, 25, 30);
            logger.Info("first");
            dateTimeForTests.Now = new DateTime(2020, 6, 30, 7, 25, 31);
            logger.Error("second");
            logger.Clear();
            dateTimeForTests.Now = new DateTime(2020, 6, 30, 7, 25, 32);
            logger.Info("third");

            var logAsRichText = logger.GetLogAsRichText(false);

            Assert.AreEqual(
                "{\\rtf1{\\colortbl;\\red0\\green0\\blue0;}\r\n" +
                "\\cf1 6/30/2020 7:25:32 AM: \\cf1 third\\par\r\n",
                logAsRichText);
        }
    }
}
