using System;
using System.IO;
using System.Runtime.Serialization.Formatters.Binary;
using System.Collections.Generic;
using Excel = Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;

namespace salesApplication
{
    public class SaleItems
    {
        private string categories;
        private string items;
        private int quantities;
        private decimal prices;

        public SaleItems (string c, string i, int q, decimal p)
        {
            this.categories = c;
            this.items = i;
            this.quantities = q;
            this.prices = p;
        }

        public string displayCategories()
        {
            return this.categories;
        }

        public string displayItems()
        {
            return this.items;
        }

        public int displayQuantities()
        {
            return this.quantities;
        }

        public decimal displayPrices()
        {
            return this.prices;
        }

        

    }

    class Program
    {
        const int textFileDataSegments = 4; //the number of segments the text file has to read from
        const int startIndex = 0; //the index number to start the serial count for all of the items used to create the objects
        const int numberOfItems = 21; //the programmer needs to know the number of number of available items for sale at all times
        const string optionOne = "y";
        const string optionTwo = "n";

        
        static void Main(string[] args)
        {
            Console.BackgroundColor = ConsoleColor.DarkBlue; //code to change the color of the background color for the text
            Console.WriteLine("Hello! Welcome to My Online Store." +
                "\n\nWhat will you like to shop for?\n");
            Console.BackgroundColor = ConsoleColor.Black;

            StreamReader salesFile = new StreamReader("salesFileTest.txt");

            int numLines = File.ReadAllLines("salesFileTest.txt").Length;

            List<string> category = new List<string>();
            List<string> item = new List<string>();
            List<int> quantity = new List<int>();
            List<decimal> price = new List<decimal>();

            for (int i = 0; i < (numLines / textFileDataSegments); ++i) 
            {
                category.Add(salesFile.ReadLine());
                item.Add(salesFile.ReadLine());
                quantity.Add(int.Parse(salesFile.ReadLine()));
                price.Add(decimal.Parse(salesFile.ReadLine()));
            }

            salesFile.Close();

            /*Excel.Application salesFile = new Excel.Application();
            Excel.Workbook salesWorkBook = salesFile.Workbooks.Open("salesFile.xlxs");
            Excel.Worksheet salesWorkSheet = (Excel.Worksheet)salesWorkBook.Worksheets.get_Item(1);

            Excel.Range salesRange = salesWorkSheet.UsedRange;
            int totalRows = salesRange.Rows.Count;
            int totalColumns = salesRange.Columns.Count;


            string firstValue, secondValue, thirdValue, fourthValue;

            for (int rowCount = 2; rowCount <= totalRows; rowCount++)
            {
                firstValue = Convert.ToString((salesRange.Cells[rowCount, 1] as Excel.Range).Text);
                secondValue = Convert.ToString((salesRange.Cells[rowCount, 2] as Excel.Range).Text);
                thirdValue = Convert.ToString((salesRange.Cells[rowCount, 3] as Excel.Range).Text);
                fourthValue = Convert.ToString((salesRange.Cells[rowCount, 4] as Excel.Range).Text);

                Console.WriteLine(firstValue + "\t" + secondValue + "\t" + thirdValue + "\t" + fourthValue);

            }

            salesWorkBook.Close();
            salesFile.Quit();

            Marshal.ReleaseComObject(salesWorkSheet);
            Marshal.ReleaseComObject(salesWorkBook);
            Marshal.ReleaseComObject(salesFile); */

            List<SaleItems> items = new List<SaleItems>();

            for (int i  = 0; i < (numLines / textFileDataSegments); ++i)
            {
                items.Add(new SaleItems(category[i], item[i], quantity[i], price[i]));
            }

            string[] categoriesArray = new string[numLines / textFileDataSegments];

            int index = startIndex;
            
            foreach (SaleItems categoriesList in items)
            {
                categoriesArray[index] = categoriesList.displayCategories();
                ++index;
            }

            List<SaleItems> shopperBag = new List<SaleItems>();
            string continueShopping = "";
            string proceedToCheckOut = "";

            List<string> booksList = itemsList("Books", items);
            List<string> laptopsPCsList = itemsList("Laptops & PCs", items);
            List<string> smartPhonesList = itemsList("Smart Phones", items);
            List<string> stationaryList = itemsList("Stationary", items);

            List<int> booksQuantity = quantityList("Books", items);
            List<int> laptopsPCsQuantity = quantityList("Laptops & PCs", items);
            List<int> smartPhonesQuantity = quantityList("Smart Phones", items);
            List<int> stationaryQuantity = quantityList("Stationary", items);

            List<decimal> booksPrice = priceList("Books", items);
            List<decimal> laptopsPCsPrice = priceList("Laptops & PCs", items);
            List<decimal> smartPhonesPrice = priceList("Smart Phones", items);
            List<decimal> stationaryPrice = priceList("Stationary", items);

            do
            {
                do
                {
                    arrayWithoutDuplicates(categoriesArray);

                    int categorySelection = userInput("\nInput the number of the category you want to browse through: ", 1, 4);

                    if (categorySelection == 1)
                    {
                        Console.WriteLine("\nYou have selected " + categorySelection);

                        availableItems("Books", items);

                        int itemSelection = userInput("\nInput the serial number for the item you want: ", 1, booksList.Count);

                        if (booksQuantity[itemSelection - 1] != 0)
                        {
                            int selectedQuantity = quantitySelection(booksQuantity, booksList, itemSelection);

                            shopperBag.Add(new SaleItems("Books", booksList[itemSelection - 1], selectedQuantity, booksPrice[itemSelection - 1]));

                            updateQuantity(booksQuantity, itemSelection, selectedQuantity);
                        }

                        else
                        {
                            Console.WriteLine("Sorry! This item is out of stock.");
                        }

                        continueShopping = userResponse("\nContinue shopping? Enter 'y' or 'n': ", optionOne, optionTwo);
                    }

                    else if (categorySelection == 2)
                    {
                        Console.WriteLine("\nYou have selected " + categorySelection);

                        availableItems("Laptops & PCs", items);

                        int itemSelection = userInput("\nInput the serial number for the item you want: ", 1, laptopsPCsList.Count);

                        if (laptopsPCsQuantity[itemSelection - 1] != 0)
                        {
                            int selectedQuantity = quantitySelection(laptopsPCsQuantity, laptopsPCsList, itemSelection);

                            shopperBag.Add(new SaleItems("Laptops & PCs", laptopsPCsList[itemSelection - 1], selectedQuantity, laptopsPCsPrice[itemSelection - 1]));

                            updateQuantity(laptopsPCsQuantity, itemSelection, selectedQuantity);
                        }

                        else
                        {
                            Console.WriteLine("Sorry! This item is out of stock.");
                        }

                        continueShopping = userResponse("\nContinue shopping? Enter 'y' or 'n': ", optionOne, optionTwo);
                    }

                    else if (categorySelection == 3)
                    {
                        Console.WriteLine("\nYou have selected " + categorySelection);
                        availableItems("Smart Phones", items);

                        int itemSelection = userInput("\nInput the serial number for the item you want: ", 1, smartPhonesList.Count);

                        if (smartPhonesQuantity[itemSelection - 1] != 0)
                        {
                            int selectedQuantity = quantitySelection(smartPhonesQuantity, smartPhonesList, itemSelection);

                            shopperBag.Add(new SaleItems("Smart Phones", smartPhonesList[itemSelection - 1], selectedQuantity, smartPhonesPrice[itemSelection - 1]));

                            updateQuantity(smartPhonesQuantity, itemSelection, selectedQuantity);
                        }

                        else
                        {
                            Console.WriteLine("Sorry! This item is out of stock.");
                        }

                        continueShopping = userResponse("\nContinue shopping? Enter 'y' or 'n': ", optionOne, optionTwo);
                    }

                    else if (categorySelection == 4)
                    {
                        Console.WriteLine("\nYou have selected " + categorySelection);
                        availableItems("Stationary", items);
                        int itemSelection = userInput("\nInput the serial number for the item you want: ", 1, stationaryList.Count);

                        if (stationaryQuantity[itemSelection - 1] != 0)
                        {
                            int selectedQuantity = quantitySelection(stationaryQuantity, stationaryList, itemSelection);

                            shopperBag.Add(new SaleItems("Stationary", stationaryList[itemSelection - 1], selectedQuantity, stationaryPrice[itemSelection - 1]));

                            updateQuantity(stationaryQuantity, itemSelection, selectedQuantity);
                        }

                        else
                        {
                            Console.WriteLine("Sorry! This item is out of stock.");
                        }

                        continueShopping = userResponse("\nContinue shopping? Enter 'y' or 'n': ", optionOne, optionTwo);
                    }

                }
                while (continueShopping == optionOne);

                Console.WriteLine("\nThank you for shopping with us! :) ");

                Console.BackgroundColor = ConsoleColor.DarkBlue;
                Console.WriteLine("\nYour selected items are: ");
                Console.WriteLine("\nS/N. ITEMS\t QUANTITY \t PRICE(GBP)");
                Console.BackgroundColor = ConsoleColor.Black;
                Console.WriteLine(" ");

                int serialNumber = startIndex;
                for (int i = 0; i < shopperBag.Count; ++i)
                {
                    Console.WriteLine(++serialNumber + ". " + shopperBag[i].displayItems() + "\t" + shopperBag[i].displayQuantities() + "\t" + shopperBag[i].displayPrices());
                }

                string deleteItem = userResponse("\nRemove an item from your bag? Enter 'y' or 'n': ", optionOne, optionTwo);

                while (deleteItem == optionOne)
                {
                    int itemSelection = userInput("\nInput the serial number of the item to remove: ", 1, shopperBag.Count);

                    if (shopperBag[itemSelection - 1].displayCategories() == "Books")
                    {
                        for (int i = 0; i < booksList.Count; ++i)
                        {
                            if (booksList[i] == shopperBag[itemSelection - 1].displayItems())
                            {
                                booksQuantity[i] = shopperBag[itemSelection - 1].displayQuantities() + booksQuantity[i];
                            }
                        }
                    }

                    else if (shopperBag[itemSelection - 1].displayCategories() == "Laptops & PCs")
                    {
                        for (int i = 0; i < laptopsPCsList.Count; ++i)
                        {
                            if (laptopsPCsList[i] == shopperBag[itemSelection - 1].displayItems())
                            {
                                laptopsPCsQuantity[i] = shopperBag[itemSelection - 1].displayQuantities() + laptopsPCsQuantity[i];
                            }
                        }
                    }

                    else if (shopperBag[itemSelection - 1].displayCategories() == "Stationary")
                    {
                        for (int i = 0; i < stationaryList.Count; ++i)
                        {
                            if (stationaryList[i] == shopperBag[itemSelection - 1].displayItems())
                            {
                                stationaryQuantity[i] = shopperBag[itemSelection - 1].displayQuantities() + stationaryQuantity[i];
                            }
                        }
                    }

                    else if (shopperBag[itemSelection - 1].displayCategories() == "Smart Phones")
                    {
                        for (int i = 0; i < smartPhonesList.Count; ++i)
                        {
                            if (smartPhonesList[i] == shopperBag[itemSelection - 1].displayItems())
                            {
                                smartPhonesQuantity[i] = shopperBag[itemSelection - 1].displayQuantities() + smartPhonesQuantity[i];
                            }
                        }
                    }

                    shopperBag.RemoveAt(itemSelection - 1);

                    Console.WriteLine("\n" + shopperBag[itemSelection - 1].displayItems() + " has been removed");

                    Console.BackgroundColor = ConsoleColor.DarkBlue;
                    Console.WriteLine("\nYour selected items are: ");
                    Console.WriteLine("\nS/N. ITEMS\t QUANTITY \t PRICE(GBP)");
                    Console.BackgroundColor = ConsoleColor.Black;
                    Console.WriteLine(" ");

                    serialNumber = startIndex;
                    for (int i = 0; i < shopperBag.Count; ++i)
                    {
                        Console.WriteLine(++serialNumber + ". " + shopperBag[i].displayItems() + "\t" + shopperBag[i].displayQuantities() + "\t" + shopperBag[i].displayPrices());
                    }

                    if (shopperBag.Count != 0)
                    {
                        deleteItem = userResponse("\nRemove an item from your bag? Enter 'y' or 'n': ", optionOne, optionTwo);
                    }

                    else
                    {
                        Console.WriteLine("Your shopping bag is empty");
                        break;
                    }
                }

                proceedToCheckOut = userResponse("\nProceed to checkout? Enter 'y' or 'n': ", optionOne, optionTwo);
            }

            while (proceedToCheckOut == optionTwo);
            

            decimal totalBill = 0;

            decimal[] billArray = new decimal[shopperBag.Count]; 
            for(int i = 0; i < shopperBag.Count; ++i)
            {
                billArray[i] = shopperBag[i].displayQuantities() * shopperBag[i].displayPrices();
            }
            for(int i = 0; i < billArray.Length; ++i)
            {
                totalBill += billArray[i];
            }

            Console.BackgroundColor = ConsoleColor.Red;
            Console.WriteLine("\nYour bill is: £" + totalBill);
            Console.BackgroundColor = ConsoleColor.Black;
        }

        static int userInput (string prompt, int lowerRange, int higherRange)
        {
            do
            {
                try
                {
                    int userSelection;

                    Console.Write(prompt);
                    userSelection = int.Parse(Console.ReadLine());

                    while (userSelection < lowerRange || userSelection > higherRange)
                    {
                        Console.Write("Invalid integer input!" + prompt);
                        userSelection = int.Parse(Console.ReadLine());

                    }

                    return userSelection;
                }
                catch
                {
                    Console.Write("Invalid integer input! Please enter an integer value.");
                }
            }
            while (true);
            
        }

        static string userResponse(string prompt, string option1, string option2) //this function makes it possible for you to ensure that the user is prompted to put in another string input if they do not put in a value but press enter just like that
        {
            
            do
            {
                try
                {
                    string response;
                    Console.Write(prompt);
                    response = Console.ReadLine();

                    while(response != option1 && response != option2)
                    {
                        Console.Write("Invalid response! " + prompt);
                        response = Console.ReadLine();
                    }

                    return response;
                }

                catch
                {
                    Console.Write("Invalid response! " + prompt);
                }
            }
            while (true);

        }

        static int updateQuantity(List<int> myList, int itemSelection, int quantitySelection)
        {
            int newValue = myList[itemSelection - 1] - quantitySelection;
            int realQuantity = myList[itemSelection - 1];

            for (int i = 0; i < myList.Count; ++i)
            {
                if (myList[i] == realQuantity)
                {
                    myList[i] = newValue;
                }
            }
            return newValue;
        }

        static void arrayWithoutDuplicates(string [] myArray)
        {
            Array.Sort(myArray);
            int j = startIndex;

            for (int i = 0; (i < myArray.Length - 1); i++)
            {
                if (myArray[i] != myArray[i + 1])
                {
                    myArray[j++] = myArray[i];
                }
            }
            myArray[j++] = myArray[myArray.Length - 1];

            int serialNumber = startIndex;
            for (int i = 0; i < j; ++i)
            {
                Console.WriteLine(++serialNumber + ". " + myArray[i]);
            }
        }

        static void availableItems(string category, List <SaleItems> myList)
        {

            Console.BackgroundColor = ConsoleColor.DarkBlue;
            Console.WriteLine("\nS/N. ITEMS \t PRICE(GBP)");
            Console.BackgroundColor = ConsoleColor.Black;
            Console.WriteLine("\n");
            int serialNumber = startIndex;
            for (int i = 0; i < numberOfItems; ++i)
            {
                if (myList[i].displayCategories() == category)
                {
                    Console.WriteLine(++serialNumber + ". " + myList[i].displayItems() + "\t" + myList[i].displayPrices());
                }
            }
        }

        static List <string> itemsList (string category, List<SaleItems> myList)
        {
            List<string> stringList = new List<string>();

            for (int i = 0; i < numberOfItems; ++i)
            {
                if (myList[i].displayCategories() == category)
                {
                    stringList.Add(myList[i].displayItems());
                }
            }
            return stringList;
        }

        static List<int> quantityList(string category, List<SaleItems> myList)
        {
            List<int> intList = new List<int>();

            for (int i = 0; i < numberOfItems; ++i)
            {
                if (myList[i].displayCategories() == category)
                {
                    intList.Add(myList[i].displayQuantities());
                }
            }

            return intList;
        }

        static List<decimal> priceList(string category, List<SaleItems> myList)
        {
            List<decimal> priceList = new List<decimal>();

            for (int i = 0; i < numberOfItems; ++i)
            {
                if (myList[i].displayCategories() == category)
                {
                    priceList.Add(myList[i].displayPrices());
                }
            }
            return priceList;
        }

        static int quantitySelection(List<int> quantityList, List<string> itemList, int itemSelection)
        {
            Console.BackgroundColor = ConsoleColor.DarkBlue;
            Console.WriteLine("\nThere are " + quantityList[itemSelection - 1] + " available " + itemList[itemSelection - 1]);
            Console.BackgroundColor = ConsoleColor.Black;

            int quantitySelection = userInput("\nInput the quantity for the item you want: ", 1, quantityList[itemSelection - 1]);

            Console.BackgroundColor = ConsoleColor.DarkBlue;
            Console.WriteLine("\n" + itemList[itemSelection - 1] + " X " + quantitySelection + " has been added to your shooping bag");
            Console.BackgroundColor = ConsoleColor.Black;

            return quantitySelection;

        }


    }
}
