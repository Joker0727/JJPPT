﻿using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Firefox;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace JJPPT
{
   public class SeleniumHelper
    {
        public IWebDriver driver = null;

        public SeleniumHelper(int driverType = 0)
        {
            switch (driverType)
            {
                case 0:
                    {
                        driver = new FirefoxDriver();
                        break;
                    }
                case 1:
                    {
                        driver = new ChromeDriver();
                        break;
                    }
                default:
                    break;
            }
           
            WinTools wt = new WinTools();
            wt.MiniMizeAppication("geckodriver");
            wt.MiniMizeAppication("chromedriver"); 
        }

        /// <summary>
        /// 根据内容选择选择下拉框
        /// </summary>
        /// <param name="select">select节点对象</param>
        /// <param name="text">option的value匹配内容</param>
        public void SelectUsage(IWebElement select, string text)
        {
            try
            {
                IList<IWebElement> AllOptions = select.FindElements(By.TagName("option"));
                foreach (IWebElement option in AllOptions)
                {
                    if (option.GetAttribute("value").Equals(text))
                    {
                        option.Click();
                        break;
                    }
                }
            }
            catch (Exception ex)
            {

            }
        }
        /// <summary>
        /// 根据id查找元素
        /// </summary>
        /// <param name="id"></param>
        /// <returns></returns>
        public IWebElement FindElementById(string id)
        {
            IWebElement element = null;
            try
            {
                element = driver.FindElement(By.Id(id));
            }
            catch (Exception ex)
            {

            }
            return element;
        }
        /// <summary>
        /// 根据name查找元素
        /// </summary>
        /// <param name="name"></param>
        /// <returns></returns>
        public IWebElement FindElementByClassName(string className)
        {
            IWebElement element = null;
            try
            {
                element = driver.FindElement(By.ClassName(className));
            }
            catch (Exception ex)
            {

            }
            return element;
        }
        /// <summary>
        /// 根据classname查找到多个元素
        /// </summary>
        /// <param name="className"></param>
        /// <returns></returns>
        public ReadOnlyCollection<IWebElement> FindElementsByClassName(string className)
        {
            ReadOnlyCollection<IWebElement> elements = null;
            try
            {
                elements = driver.FindElements(By.ClassName(className));
            }
            catch (Exception ex)
            {

            }
            return elements;
        }
        /// <summary>
        /// 根据name查找元素
        /// </summary>
        /// <param name="name"></param>
        /// <returns></returns>
        public IWebElement FindElementByName(string name)
        {
            IWebElement element = null;
            try
            {
                element = driver.FindElement(By.Name(name));
            }
            catch (Exception ex)
            {

            }
            return element;
        }
        /// <summary>
        /// 根据页面链接文字查找文字
        /// </summary>
        /// <param name="text"></param>
        /// <returns></returns>
        public IWebElement FindElementByLinkText(string text)
        {
            IWebElement element = null;
            try
            {
                element = driver.FindElement(By.LinkText(text));
            }
            catch (Exception ex)
            {

            }
            return element;
        }
        /// <summary>
        /// 根据CSS查找一组元素集合
        /// </summary>
        /// <param name="cssString">.classNameA.classNameB</param>
        /// <returns></returns>
        public ReadOnlyCollection<IWebElement> FindElementsByCss(string cssString)
        {
            ReadOnlyCollection<IWebElement> elements = null;
            try
            {
                elements = driver.FindElements(By.CssSelector(cssString));
            }
            catch (Exception ex)
            {

            }
            return elements;
        }
        /// <summary>
        /// 根据CSS查找一个
        /// </summary>
        /// <param name="cssString">.classNameA.classNameB</param>
        /// <returns></returns>
        public IWebElement FindElementByCss(string cssString)
        {
            IWebElement element = null;
            try
            {
                element = driver.FindElement(By.CssSelector(cssString));
            }
            catch (Exception ex)
            {

            }
            return element;
        }
        /// <summary>
        /// 根据XPath查找元素
        /// </summary>
        /// <param name="xpath"></param>
        /// <returns></returns>
        public IWebElement FindElementByXPath(string xpath)
        {
            IWebElement element = null;
            try
            {
                element = driver.FindElement(By.XPath(xpath));
            }
            catch (Exception ex)
            {

            }
            return element;
        }
        /// <summary>
        /// 根据XPath查找元素，返回集合
        /// </summary>
        /// <param name="xpath"></param>
        /// <returns></returns>
        public ReadOnlyCollection<IWebElement> FindElementsByXPath(string xpath)
        {
            ReadOnlyCollection<IWebElement> elements = null;
            try
            {
                elements = driver.FindElements(By.XPath(xpath));
            }
            catch (Exception ex)
            {

            }
            return elements;
        }
    }
}
