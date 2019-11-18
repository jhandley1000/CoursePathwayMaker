using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Remote;

namespace CoursePathwayMaker.NuStarDataScraperTool
{
    public class NuStarWebsiteNavigator
    {
        ChromeDriver driver;
        const string url = "https://nustar.newcastle.edu.au/psp/CS92PRD/EMPLOYEE/SA/c/ESTABLISH_COURSES.CLASS_ROSTER.GBL?PORTALPARAM_PTCNAV=HC_CLASS_ROSTER_GBL&EOPP.SCNode=HRMS&EOPP.SCPortal=EMPLOYEE&EOPP.SCName=HCSR_CURRICULUM_MANAGEMENT&EOPP.SCLabel=Curriculum%20Management&EOPP.SCPTfname=HCSR_CURRICULUM_MANAGEMENT&FolderPath=PORTAL_ROOT_OBJECT.HCSR_CURRICULUM_MANAGEMENT.HCSR_CLASS_ROSTER.HC_CLASS_ROSTER_GBL&IsFolder=false";
        const string altUrl = "https://nustar.newcastle.edu.au/psc/CS92PRD/EMPLOYEE/SA/c/ESTABLISH_COURSES.CLASS_ROSTER.GBL?PORTALPARAM_PTCNAV=HC_CLASS_ROSTER_GBL&EOPP.SCNode=HRMS&EOPP.SCPortal=EMPLOYEE&EOPP.SCName=HCSR_CURRICULUM_MANAGEMENT&EOPP.SCLabel=Curriculum%20Management&EOPP.SCPTfname=HCSR_CURRICULUM_MANAGEMENT&FolderPath=PORTAL_ROOT_OBJECT.HCSR_CURRICULUM_MANAGEMENT.HCSR_CLASS_ROSTER.HC_CLASS_ROSTER_GBL&IsFolder=false&PortalActualURL=https%3a%2f%2fnustar.newcastle.edu.au%2fpsc%2fCS92PRD%2fEMPLOYEE%2fSA%2fc%2fESTABLISH_COURSES.CLASS_ROSTER.GBL&PortalContentURL=https%3a%2f%2fnustar.newcastle.edu.au%2fpsc%2fCS92PRD%2fEMPLOYEE%2fSA%2fc%2fESTABLISH_COURSES.CLASS_ROSTER.GBL&PortalContentProvider=SA&PortalCRefLabel=Class%20Roster&PortalRegistryName=EMPLOYEE&PortalServletURI=https%3a%2f%2fnustar.newcastle.edu.au%2fpsp%2fCS92PRD%2f&PortalURI=https%3a%2f%2fnustar.newcastle.edu.au%2fpsc%2fCS92PRD%2f&PortalHostNode=HRMS&NoCrumbs=yes&PortalKeyStruct=yes";

        public NuStarWebsiteNavigator(ChromeDriver driver)
        {
            this.driver = driver;
        }

        public void NavigateToUonWebsite()
        {
            driver.Navigate().GoToUrl(altUrl);
        }

        public void FillSearchFilters(string subjectArea, string term)
        {
            fillSearchFilter("CLASS_RSTR_SRCH_INSTITUTION", "UNAUS");
            fillSearchFilter("CLASS_RSTR_SRCH_STRM", term);
            fillSearchFilter("CLASS_RSTR_SRCH_SUBJECT", subjectArea);

            ClickSearchButton();
        }


        public void ReturnToClassRosterSearch()
        {
            bool staleElement = true;
            var loopCount = 0;
            while (staleElement && loopCount < 3)
            {
                try
                {
                    driver.FindElement(By.Id("#ICList")).Click();
                    staleElement = false;

                }
                catch (NoSuchElementException e)
                {
                    try
                    {
                        driver.FindElementById("ptifrmtgtframe").Submit();
                    }
                    catch { }
                    staleElement = true;
                    loopCount++;

                    if (loopCount > 4)
                    {
                        throw;
                    }
                }
            }
        }

        public void GoToSearchResultEnrollmentPage(IWebElement tableRow)
        {
            tableRow.FindElements(By.TagName("a")).Last().Click();
        }

        void fillSearchFilter(string fieldID, string fillString)
        {
            var inputBox = driver.FindElement(By.Id(fieldID));
            inputBox.SendKeys(fillString);
        }

        void ClickSearchButton()
        {
            driver.FindElementById("#ICSearch").Click();
        }
    }
}
