#include <iostream>
#include <conio.h>
#include <windows.h>
#include <sstream>
#include <regex>
#include <fstream>
int main();
using namespace std;
HANDLE hConsole = GetStdHandle(STD_OUTPUT_HANDLE);
int ArrCount = 0;
string Arr[10000000];
int rowCount;
int columnCount;
bool isInteger(string s)
{
    istringstream iss(s);
    int n;
    iss >> noskipws >> n;
    return iss.eof() && !iss.fail();
}

bool isFloat(string s)
{
    istringstream iss(s);
    float f;
    iss >> noskipws >> f;
    return iss.eof() && !iss.fail();
}
string Spaces(string s)
{
    regex pattern("\\s+$");
    return regex_replace(s, pattern, "");
}
enum Color
{
    Aqua,
    Purple,
    Yellow
};
enum DataType
{
    Int,
    Float,
    String
};
void gotoxy(int x, int y)
{
    COORD coordinates;
    coordinates.X = x;
    coordinates.Y = y;
    SetConsoleCursorPosition(GetStdHandle(STD_OUTPUT_HANDLE), coordinates);
}
class Cell
{
    friend class node;

private:
    int x;
    int y;
    string data = "    ";

public:
    Color color;
    DataType dataType;
    Cell()
    {
        x = y = 0;
        color = Yellow;
        data = "    ";
    }
    string getData()
    {
        return data;
    }
    void setData()
    {
        data = "    ";
    }
    int getX()
    {
        return x;
    }
    int getY()
    {
        return y;
    }
    void setDataType(int val)
    {
        if (val == 1)
            dataType = Int;
        else if (val == 2)
            dataType = Float;
        else
            dataType = String;
    }
    DataType getDataType()
    {
        return dataType;
    }
    void deselect()
    {
        color = Yellow;
    }
    void select()
    {
        color = Purple;
    }
    int getCode()
    {
        if (color == Aqua)
        {
            return 3;
        }
        else if (color == Purple)
        {
            return 5;
        }
        else
        {
            return 6;
        }
    }
    Cell(int xCoor, int yCoor, string value)
    {
        x = xCoor;
        y = yCoor;
        data = value;
        color = Yellow;
        dataType = String;
    }
    void setData(string data)
    {
        if (isInteger(data))
            setDataType(1);
        else if (isFloat(data))
            setDataType(2);
        else
            setDataType(3);
        string s = "";
        for (int i = 0; i < 4 && i < data.length(); i++)
        {
            s += data[i];
        }
        if (data.length() < 4)
        {
            for (int i = 0; i < 4 - data.length(); i++)
            {
                s += " ";
            }
        }
        this->data = s;
    }
};
class node
{
    friend class Excel;
    friend class Cell;

public:
    Cell *data;
    node *top;
    node *bottom;
    node *left;
    node *right;
    node()
    {
        data = new Cell();
        top = nullptr;
        bottom = nullptr;
        left = nullptr;
        right = nullptr;
    }
    node(Cell *value)
    {
        data = value;
        top = nullptr;
        bottom = nullptr;
        left = nullptr;
        right = nullptr;
    }
    void location()
    {
        node *newNode = this;
        int counter = 0;
        while (newNode->top != nullptr)
        {
            counter++;
            newNode = newNode->top;
        }
        data->y = counter;
        counter = 0;
        while (newNode->left != nullptr)
        {
            counter++;
            newNode = newNode->left;
        }
        data->x = counter;
    }
};
class iterator
{
    node *iter;

public:
    iterator()
    {
        iter = nullptr;
    }
    iterator(node *n)
    {
        iter = n;
    }
    iterator operator++()
    {
        if (iter->bottom != nullptr)
            iter = iter->bottom;
        return *this;
    }
    iterator operator--()
    {
        if (iter->top != nullptr)
            iter = iter->top;
        return *this;
    }
    iterator operator++(int)
    {
        if (iter->right != nullptr)
            iter = iter->right;
        return *this;
    }
    iterator operator--(int)
    {
        if (iter->left != nullptr)
            iter = iter->left;
        return *this;
    }
    bool operator==(iterator i)
    {
        return (iter == i.iter);
    }
    bool operator!=(iterator i)
    {
        return (iter != i.iter);
    }

    friend class Excel;
};
class Excel
{
    friend class Cell;
    friend class node;

private:
    node *selectedNode;

    void updatePrevCell(node *prev)
    {
        SetConsoleTextAttribute(hConsole, prev->data->getCode());
        gotoxy((prev->data->getX() * 6), (prev->data->getY() * 4));
        cout << " ____ " << endl;
        gotoxy((prev->data->getX() * 6), (prev->data->getY() * 4) + 1);
        cout << "|    |" << endl;
        gotoxy((prev->data->getX() * 6), (prev->data->getY() * 4) + 2);
        cout << "|" << prev->data->getData() << "|";
        gotoxy((prev->data->getX() * 6), (prev->data->getY() * 4) + 3);
        cout << "|    |";
        gotoxy(prev->data->getX() * 6, (prev->data->getY() * 4) + 3);
        cout << "|____|";
    }
    void updateSelectedCell()
    {
        SetConsoleTextAttribute(hConsole, selectedNode->data->getCode());
        gotoxy((selectedNode->data->getX() * 6), (selectedNode->data->getY() * 4));
        cout << " ____ " << endl;
        gotoxy((selectedNode->data->getX() * 6), (selectedNode->data->getY() * 4) + 1);
        cout << "|    |" << endl;
        gotoxy((selectedNode->data->getX() * 6), (selectedNode->data->getY() * 4) + 2);
        cout << "|" << selectedNode->data->getData() << "|";
        gotoxy((selectedNode->data->getX() * 6), (selectedNode->data->getY() * 4) + 3);
        cout << "|    |";
        gotoxy(selectedNode->data->getX() * 6, (selectedNode->data->getY() * 4) + 3);
        cout << "|____|";
    }

public:
    Excel(int rowCount, int columnCount)
    {
        selectedNode = new node();
        for (int i = 0; i < rowCount - 1; i++)
        {
            extendRow();
        }
        for (int i = 0; i < columnCount - 1; i++)
        {
            extendColumn();
        }
    }
    node *getTop()
    {
        node *temp = selectedNode;
        while (temp->top)
        {
            temp = temp->top;
        }
        return temp;
    }
    node *getLeft()
    {
        node *temp = selectedNode;
        while (temp->left)
        {
            temp = temp->left;
        }
        return temp;
    }
    node *getNode()
    {
        return selectedNode;
    }
    node *getTopRight()
    {
        node *temp = selectedNode;
        while (temp->top)
        {
            temp = temp->top;
        }
        while (temp->right)
        {
            temp = temp->right;
        }
        return temp;
    }
    node *getTopLeft()
    {
        node *temp = selectedNode;
        while (temp->top)
        {
            temp = temp->top;
        }
        while (temp->left)
        {
            temp = temp->left;
        }
        return temp;
    }
    node *getBottomLeft()
    {
        node *temp = selectedNode;
        while (temp->left)
        {
            temp = temp->left;
        }
        while (temp->bottom)
        {
            temp = temp->bottom;
        }
        return temp;
    }
    void moveUp()
    {
        node *temp = selectedNode;
        if (selectedNode->top != nullptr)
        {
            selectedNode->data->deselect();
            selectedNode = selectedNode->top;
            selectedNode->data->select();
            updateSheet(temp);
        }
    }
    void moveDown()
    {
        node *temp = selectedNode;
        if (selectedNode->bottom != nullptr)
        {
            selectedNode->data->deselect();
            selectedNode = selectedNode->bottom;
            selectedNode->data->select();
            updateSheet(temp);
        }
        else
        {
            rowCount++;
            extendRow();
        }
    }
    void moveLeft()
    {
        node *temp = selectedNode;
        if (selectedNode->left != nullptr)
        {
            selectedNode->data->deselect();
            selectedNode = selectedNode->left;
            selectedNode->data->select();
            updateSheet(temp);
        }
    }
    void moveRight()
    {
        node *temp = selectedNode;
        if (selectedNode->right != nullptr)
        {
            selectedNode->data->deselect();
            selectedNode = selectedNode->right;
            selectedNode->data->select();
            updateSheet(temp);
        }
        else
        {
            columnCount++;
            extendColumn();
        }
    }
    void extendColumn()
    {
        node *temp = getTopRight();
        while (temp)
        {
            node *newNode = new node();
            temp->right = newNode;
            temp->right->left = temp;
            temp = temp->bottom;
        }
        temp = getTopRight();
        while (temp->left->bottom)
        {
            temp->bottom = temp->left->bottom->right;
            temp->bottom->top = temp;
            temp = temp->bottom;
        }
        printSheet();
    }
    void extendColumnRight()
    {
        node *temp = getTop();
        node *newNode;
        if (temp->right == nullptr)
        {
            extendColumn();
        }
        else
        {
            while (temp != nullptr)
            {
                node *newNode = new node();
                node *temp2 = temp->right;
                temp->right = newNode;
                newNode->right = temp2;
                temp2->left = newNode;
                newNode->left = temp;
                temp = temp->bottom;
            }
            temp = getTop();
            while (temp->bottom)
            {
                temp->right->bottom = temp->bottom->right;
                temp->right->bottom->top = temp->right;
                temp = temp->bottom;
            }
        }
        columnCount++;
    }
    void extendColumnLeft()
    {
        node *temp = getTop();
        node *newNode;
        if (temp->left == nullptr)
        {
            temp = getTopLeft();
            while (temp != nullptr)
            {
                newNode = new node();
                temp->left = newNode;
                temp->left->right = temp;
                temp = temp->bottom;
            }
            temp = getTopLeft();
            while (temp->right->bottom)
            {
                temp->bottom = temp->right->bottom->left;
                temp->bottom->top = temp;
                temp = temp->bottom;
            }
        }
        else
        {
            while (temp != nullptr)
            {
                newNode = new node();
                node *temp2 = temp->left;
                temp->left = newNode;
                newNode->left = temp2;
                temp2->right = newNode;
                newNode->right = temp;
                temp = temp->bottom;
            }
            temp = getTop();
            while (temp->bottom)
            {
                temp->left->bottom = temp->bottom->left;
                temp->left->bottom->top = temp->left;
                temp = temp->bottom;
            }
        }
        columnCount++;
    }
    void extendRow()
    {
        node *temp = getBottomLeft();
        while (temp)
        {
            node *newNode = new node();
            temp->bottom = newNode;
            temp->bottom->top = temp;
            temp = temp->right;
        }
        temp = getBottomLeft();
        while (temp->top->right)
        {
            temp->right = temp->top->right->bottom;
            temp->right->left = temp;
            temp = temp->right;
        }
        printSheet();
    }
    void extendRowUp()
    {
        rowCount++;
        node *temp = getLeft();
        node *newNode;
        if (temp->top == nullptr)
        {
            temp = getTopLeft();
            while (temp)
            {
                newNode = new node();
                temp->top = newNode;
                newNode->bottom = temp;

                if (temp->left)
                {
                    newNode->left = temp->left;
                    temp->left->right = newNode;
                }

                if (temp->right)
                {
                    newNode->right = temp->right;
                    temp->right->left = newNode;
                }

                temp = temp->right;
            }
            temp = getTopLeft();
            while (temp->bottom->right)
            {
                temp->right = temp->bottom->right->top;
                temp->right->left = temp;
                temp = temp->right;
            }
        }
        else
        {
            while (temp != nullptr)
            {
                newNode = new node();
                node *temp2 = temp->top;
                temp->top = newNode;
                newNode->top = temp2;
                temp2->bottom = newNode;
                newNode->bottom = temp;
                temp = temp->right;
            }
            temp = getLeft();
            while (temp->right)
            {
                temp->top->right = temp->right->top;
                temp->top->right->left = temp->top;
                temp = temp->right;
            }
        }
    }
    void extendRowBelow()
    {
        node *temp = getLeft();
        node *newNode;
        if (temp->bottom == nullptr)
        {
            extendRow();
        }
        else
        {
            while (temp != nullptr)
            {
                newNode = new node();
                node *temp2 = temp->bottom;
                temp->bottom = newNode;
                newNode->bottom = temp2;
                temp2->top = newNode;
                newNode->top = temp;
                temp = temp->right;
            }
            temp = getLeft();
            while (temp->right)
            {
                temp->bottom->right = temp->right->bottom;
                temp->bottom->right->left = temp->bottom;
                temp = temp->right;
            }
        }
        rowCount++;
    }
    void selectedCell()
    {
        node *head = selectedNode;
        head->location();
        SetConsoleTextAttribute(hConsole, head->data->getCode());
        gotoxy((head->data->getX() * 6), (head->data->getY() * 4));
        cout << "______" << endl;
        gotoxy((head->data->getX() * 6), (head->data->getY() * 4) + 1);
        cout << "|    |" << endl;
        gotoxy((head->data->getX() * 6), (head->data->getY() * 4) + 2);
        cout << "|" << head->data->getData() << "|";
        gotoxy((head->data->getX() * 6), (head->data->getY() * 4) + 3);
        cout << "|    |";
        if (!head->bottom)
        {
            cout << "______";
            gotoxy(head->data->getX() * 6, (head->data->getY() * 4) + 4);
        }
        head = head->right;
    }
    void printSheet()
    {
        system("cls");
        node *temp = getTopLeft();
        while (temp)
        {
            node *head = temp;
            while (head)
            {
                head->location();
                SetConsoleTextAttribute(hConsole, head->data->getCode());
                gotoxy((head->data->getX() * 6), (head->data->getY() * 4));
                cout << " ____ " << endl;
                gotoxy((head->data->getX() * 6), (head->data->getY() * 4) + 1);
                cout << "|    |" << endl;
                gotoxy((head->data->getX() * 6), (head->data->getY() * 4) + 2);
                cout << "|" << head->data->getData() << "|";
                gotoxy((head->data->getX() * 6), (head->data->getY() * 4) + 3);
                cout << "|____|";
                if (!head->bottom)
                {
                    gotoxy(head->data->getX() * 6, (head->data->getY() * 4) + 3);
                    cout << "|____|";
                }
                head = head->right;
            }
            temp = temp->bottom;
        }
    }
    void updateSheet(node *prev)
    {
        updatePrevCell(prev);
        updateSelectedCell();
    }
    void enterData()
    {
        string abc = "";
        cout << "Enter data: ";
        cin >> abc;
        selectedNode->data->setData(abc);
        printSheet();
    }
    void deleteColumn()
    {
        columnCount--;
        node *current = getTop();
        node *previous = nullptr;
        while (current)
        {
            node *left = current->left;
            if (left)
            {
                left->right = current->right;
            }
            if (current->right)
            {
                current->right->left = left;
            }

            if (current->top)
            {
                current->top->bottom = current->bottom;
            }
            if (current->bottom)
            {
                current->bottom->top = current->top;
            }
            previous = current;
            current = current->bottom;
            delete previous;
        }
        if (selectedNode->right == nullptr)
            selectedNode = selectedNode->left;
        else
            selectedNode = selectedNode->right;
    }
    void deleteColumnFromLast()
    {
        columnCount--;
        node *current = getTopRight();
        node *previous = nullptr;
        while (current)
        {
            node *left = current->left;
            if (left)
            {
                left->right = current->right;
            }
            if (current->right)
            {
                current->right->left = left;
            }

            if (current->top)
            {
                current->top->bottom = current->bottom;
            }
            if (current->bottom)
            {
                current->bottom->top = current->top;
            }
            previous = current;
            current = current->bottom;
            delete previous;
        }
    }
    void deleteRow()
    {
        rowCount--;
        node *current = getLeft();
        node *previous = nullptr;
        while (current)
        {
            node *top = current->top;
            if (top)
            {
                top->bottom = current->bottom;
            }

            if (current->bottom)
            {
                current->bottom->top = top;
            }

            if (current->right)
            {
                current->right->left = current->left;
            }
            if (current->left)
            {
                current->left->right = current->right;
            }
            previous = current;
            current = current->right;
            delete previous;
        }
        if (selectedNode->bottom == nullptr)
            selectedNode = selectedNode->top;
        else
            selectedNode = selectedNode->bottom;
    }
    void deleteRowFromLast()
    {
        rowCount--;
        node *current = getBottomLeft();
        node *previous = nullptr;
        while (current)
        {
            node *top = current->top;
            if (top)
            {
                top->bottom = current->bottom;
            }

            if (current->bottom)
            {
                current->bottom->top = top;
            }

            if (current->right)
            {
                current->right->left = current->left;
            }
            if (current->left)
            {
                current->left->right = current->right;
            }
            previous = current;
            current = current->right;
            delete previous;
        }
    }
    void clearRow()
    {
        node *current = getLeft();
        while (current->right)
        {
            current->data->setData();
            current = current->right;
        }
    }
    void clearColumn()
    {
        node *current = getTop();
        while (current->bottom)
        {
            current->data->setData();
            current = current->bottom;
        }
    }
    void InsertByRightShift()
    {
        node *temp = selectedNode;
        if (temp->right == nullptr)
        {
            extendColumn();
            temp->right->data->setData(temp->data->getData());
            temp->data->setData();
        }
        else if (temp->right != nullptr && temp->right->data->getData() == "    ")
        {
            temp->right->data->setData(temp->data->getData());
            temp->data->setData();
        }
        else
        {
            while (temp->right != nullptr)
            {
                temp = temp->right;
            }
            if (temp->data->getData() == "    ")
            {
                node *newNode = new node();
                while (temp != selectedNode)
                {
                    newNode = temp;
                    temp->data->setData(temp->left->data->getData());
                    temp = temp->left;
                }
                selectedNode->data->setData();
            }
            else
            {
                extendColumn();
                temp = temp->right;
                node *newNode = new node();
                while (temp != selectedNode)
                {
                    newNode = temp;
                    temp->data->setData(temp->left->data->getData());
                    temp = temp->left;
                }
                selectedNode->data->setData();
            }
        }
    }
    void InsertByDownShift()
    {
        node *temp = selectedNode;
        if (temp->bottom == nullptr)
        {
            extendRow();
            temp->bottom->data->setData(temp->data->getData());
            temp->data->setData();
        }
        else if (selectedNode->data->getData() == "    ")
            return;
        else if (temp->bottom != nullptr && temp->bottom->data->getData() == "    ")
        {
            temp->bottom->data->setData(temp->data->getData());
            temp->data->setData();
        }
        else
        {
            while (temp->bottom != nullptr)
            {
                temp = temp->bottom;
            }
            if (temp->data->getData() == "    ")
            {
                node *newNode = new node();
                while (temp != selectedNode)
                {
                    newNode = temp;
                    temp->data->setData(temp->top->data->getData());
                    temp = temp->top;
                }
                selectedNode->data->setData();
            }
            else
            {
                extendRow();
                temp = temp->bottom;
                node *newNode = new node();
                while (temp != selectedNode)
                {
                    newNode = temp;
                    temp->data->setData(temp->top->data->getData());
                    temp = temp->top;
                }
                selectedNode->data->setData();
            }
        }
    }
    void DeleteByLeftShift()
    {
        node *currentCell = selectedNode;
        if (currentCell->left != nullptr)
        {
            currentCell->data->setData();
            node *temp = currentCell->right;
            while (temp != nullptr)
            {
                currentCell->data->setData(temp->data->getData());
                currentCell = temp;
                currentCell->data->setData();
                temp = temp->right;
            }
        }
    }
    void DeleteByUpShift()
    {
        node *currentCell = selectedNode;
        if (currentCell->top != nullptr)
        {
            currentCell->data->setData();
            node *temp = currentCell->bottom;
            while (temp != nullptr)
            {
                currentCell->data->setData(temp->data->getData());
                currentCell = temp;
                currentCell->data->setData();
                temp = temp->bottom;
            }
        }
    }
    int calculateRangeSum(node *startingCell, node *endingCell)
    {
        int x = startingCell->data->getX();
        int x1 = endingCell->data->getX();
        int y = startingCell->data->getY();
        int y1 = endingCell->data->getY();
        DataType data;
        int cal = 0;
        if (x == x1)
        {
            while (startingCell != endingCell->bottom)
            {
                data = startingCell->data->getDataType();
                if (data == Int || data == Float)
                {
                    cal += stoi(Spaces(startingCell->data->getData()));
                    startingCell = startingCell->bottom;
                }
                else
                {
                    startingCell = startingCell->bottom;
                }
            }
        }
        else
        {
            while (startingCell != endingCell->right)
            {
                data = startingCell->data->getDataType();
                if (data == Int || data == Float)
                {
                    cal += stoi(Spaces(startingCell->data->getData()));
                    startingCell = startingCell->right;
                }
                else
                {
                    startingCell = startingCell->right;
                }
            }
        }
        return cal;
    }
    void calculateSum(node *startingCell, node *endingCell)
    {
        int x = startingCell->data->getX();
        int x1 = endingCell->data->getX();
        int y = startingCell->data->getY();
        int y1 = endingCell->data->getY();
        DataType data;
        int cal = 0;
        if (x == x1)
        {
            while (startingCell != endingCell->bottom)
            {
                data = startingCell->data->getDataType();
                if (data == Int || data == Float)
                {
                    cal += stoi(Spaces(startingCell->data->getData()));
                    startingCell = startingCell->bottom;
                }
                else
                {
                    startingCell = startingCell->bottom;
                }
            }
        }
        else
        {
            while (startingCell != endingCell->right)
            {
                data = startingCell->data->getDataType();
                if (data == Int || data == Float)
                {
                    cal += stoi(Spaces(startingCell->data->getData()));
                    startingCell = startingCell->right;
                }
                else
                {
                    startingCell = startingCell->right;
                }
            }
        }
        selectedNode->data->setData(to_string(cal));
    }
    float calculateRangeAverage(node *startingCell, node *endingCell)
    {
        int x = startingCell->data->getX();
        int x1 = endingCell->data->getX();
        int y = startingCell->data->getY();
        int y1 = endingCell->data->getY();
        DataType data;
        int cal = 0;
        int counter = 0;
        if (x == x1)
        {
            while (startingCell != endingCell->bottom)
            {
                data = startingCell->data->getDataType();
                if (data == Int || data == Float)
                {
                    cal += stoi(Spaces(startingCell->data->getData()));
                    startingCell = startingCell->bottom;
                    counter++;
                }
                else
                {
                    startingCell = startingCell->bottom;
                }
            }
        }
        else
        {
            while (startingCell != endingCell->right)
            {
                data = startingCell->data->getDataType();
                if (data == Int || data == Float)
                {
                    cal += stoi(Spaces(startingCell->data->getData()));
                    startingCell = startingCell->right;
                    counter++;
                }
                else
                {
                    startingCell = startingCell->right;
                }
            }
        }
        return cal / counter;
    }
    void calculateCount(node *startingCell, node *endingCell)
    {
        int x = startingCell->data->getX();
        int x1 = endingCell->data->getX();
        int y = startingCell->data->getY();
        int y1 = endingCell->data->getY();
        DataType data;
        int counter = 0;
        if (x == x1)
        {
            while (startingCell != endingCell->bottom)
            {
                data = startingCell->data->getDataType();
                if (data == Int || data == Float)
                {
                    counter++;
                    startingCell = startingCell->bottom;
                }
                else
                {
                    startingCell = startingCell->bottom;
                }
            }
        }
        else
        {
            while (startingCell != endingCell->right)
            {
                data = startingCell->data->getDataType();
                if (data == Int || data == Float)
                {
                    counter++;
                    startingCell = startingCell->right;
                }
                else
                {
                    startingCell = startingCell->right;
                }
            }
        }
        selectedNode->data->setData(to_string(counter));
    }
    void calculateAverage(node *startingCell, node *endingCell)
    {
        int x = startingCell->data->getX();
        int x1 = endingCell->data->getX();
        int y = startingCell->data->getY();
        int y1 = endingCell->data->getY();
        DataType data;
        int cal = 0;
        int counter = 0;
        if (x == x1)
        {
            while (startingCell != endingCell->bottom)
            {
                data = startingCell->data->getDataType();
                if (data == Int || data == Float)
                {
                    cal += stoi(Spaces(startingCell->data->getData()));
                    startingCell = startingCell->bottom;
                    counter++;
                }
                else
                {
                    startingCell = startingCell->bottom;
                }
            }
        }
        else
        {
            while (startingCell != endingCell->right)
            {
                data = startingCell->data->getDataType();
                if (data == Int || data == Float)
                {
                    cal += stoi(Spaces(startingCell->data->getData()));
                    startingCell = startingCell->right;
                    counter++;
                }
                else
                {
                    startingCell = startingCell->right;
                }
            }
        }
        selectedNode->data->setData(to_string(cal / counter));
    }
    void calculateMin(node *startingCell, node *endingCell)
    {
        int x = startingCell->data->getX();
        int x1 = endingCell->data->getX();
        int y = startingCell->data->getY();
        int y1 = endingCell->data->getY();
        DataType data;
        int cal = 0;
        int min;
        if (x == x1)
        {
            while (startingCell != endingCell->bottom)
            {
                data = startingCell->data->getDataType();
                if (data == Int || data == Float)
                {
                    min = stoi(Spaces(startingCell->data->getData()));
                    if (min > cal)
                    {
                        cal = min;
                    }
                    startingCell = startingCell->bottom;
                }
                else
                {
                    startingCell = startingCell->bottom;
                }
            }
        }
        else
        {
            while (startingCell != endingCell->right)
            {
                data = startingCell->data->getDataType();
                if (data == Int || data == Float)
                {
                    if (min > cal)
                    {
                        cal = min;
                    }
                    startingCell = startingCell->right;
                }
                else
                {
                    startingCell = startingCell->right;
                }
            }
        }
        selectedNode->data->setData(to_string(min));
    }
    void calculateMax(node *startingCell, node *endingCell)
    {
        int x = startingCell->data->getX();
        int x1 = endingCell->data->getX();
        int y = startingCell->data->getY();
        int y1 = endingCell->data->getY();
        DataType data;
        int cal = 0;
        int min = 0;
        if (x == x1)
        {
            while (startingCell != endingCell->bottom)
            {
                data = startingCell->data->getDataType();
                if (data == Int || data == Float)
                {
                    cal = stoi(Spaces(startingCell->data->getData()));
                    if (min < cal)
                    {
                        min = cal;
                    }
                    startingCell = startingCell->bottom;
                }
                else
                {
                    startingCell = startingCell->bottom;
                }
            }
        }
        else
        {
            while (startingCell != endingCell->right)
            {
                data = startingCell->data->getDataType();
                if (data == Int || data == Float)
                {
                    if (min < cal)
                    {
                        min = cal;
                    }
                    startingCell = startingCell->right;
                }
                else
                {
                    startingCell = startingCell->right;
                }
            }
        }
        selectedNode->data->setData(to_string(min));
    }
    node *getCell(int x, int y)
    {
        node *temp = getTopLeft();
        for (int i = 0; i < x; i++)
        {
            temp = temp->right;
        }
        for (int i = 0; i < y; i++)
        {
            temp = temp->bottom;
        }
        return temp;
    }
    void Copy(node *startingCell, node *endingCell)
    {
        ArrCount = 0;
        int x = startingCell->data->getX();
        int x1 = endingCell->data->getX();
        int y = startingCell->data->getY();
        int y1 = endingCell->data->getY();
        if (x == x1)
        {
            while (startingCell != endingCell->bottom)
            {
                Arr[ArrCount] = startingCell->data->getData();
                startingCell = startingCell->bottom;
                ArrCount++;
            }
        }
        else
        {
            while (startingCell != endingCell->right)
            {
                Arr[ArrCount] = startingCell->data->getData();
                startingCell = startingCell->right;
                ArrCount++;
            }
        }
    }
    void Cut(node *startingCell, node *endingCell)
    {
        ArrCount = 0;
        int x = startingCell->data->getX();
        int x1 = endingCell->data->getX();
        int y = startingCell->data->getY();
        int y1 = endingCell->data->getY();
        if (x == x1)
        {
            while (startingCell != endingCell->bottom)
            {
                Arr[ArrCount] = startingCell->data->getData();
                startingCell->data->setData();
                startingCell = startingCell->bottom;
                ArrCount++;
            }
        }
        else
        {
            while (startingCell != endingCell->right)
            {
                Arr[ArrCount] = startingCell->data->getData();
                startingCell->data->setData();
                startingCell = startingCell->right;
                ArrCount++;
            }
        }
    }
    void Paste(string val)
    {
        node *temp2 = selectedNode;
        if (val == "Row" || val == "row")
        {
            for (int i = 0; i < ArrCount; i++)
            {
                temp2->data->setData(Arr[i]);
                if (temp2->right == nullptr)
                {
                    extendColumn();
                }
                temp2 = temp2->right;
            }
            deleteColumnFromLast();
        }
        else
        {
            for (int i = 0; i < ArrCount; i++)
            {
                temp2->data->setData(Arr[i]);
                if (temp2->bottom == nullptr)
                {
                    extendRow();
                }
                temp2 = temp2->bottom;
            }
            deleteRowFromLast();
        }
    }
};
string parsItems(string itemName, int itemRate)
{
    int commaCount = 1;
    string item;
    for (int x = 0; x < itemName.length(); x++)
    {
        if (itemName[x] == ',')
        {
            commaCount = commaCount + 1;
        }
        else if (commaCount == itemRate)
        {
            item = item + itemName[x];
        }
    }
    return item;
}
void end()
{
    Excel excel(1, 2);
    cout << "Enter any key to continue";
    getch();
    excel.printSheet();
}
void storeInFile(int rowCount, int columnCount)
{

    fstream file;
    file.open("count.txt", ios::out);
    file << rowCount << ',' << columnCount << endl;
    file.close();
}
void loadData()
{
    fstream file;
    string word;
    file.open("count.txt", ios::in);
    string username, password, card, pinn;
    while (!file.eof())
    {
        getline(file, word);
        if (word == "")
            break;
        rowCount = stoi(parsItems(word, 1));
        columnCount = stoi(parsItems(word, 2));
    }
    file.close();
}
void saveGridToFile(string filename, Excel excel)
{
    ofstream outFile(filename);
    if (!outFile.is_open())
    {
        return;
    }
    node *temp = excel.getTopLeft();
    while (temp != nullptr)
    {
        node *head = temp;

        while (head != nullptr)
        {
            outFile << head->data->getData() << (char)200;

            head = head->right;
        }
        outFile << "\n";
        temp = temp->bottom;
    }

    outFile.close();
}

void loadGridFromFile(string filename, Excel excel)
{
    ifstream inFile(filename);
    if (!inFile.is_open())
    {
        return;
    }
    string line;
    int row = 0;
    while (getline(inFile, line))
    {
        istringstream iss(line);
        string token;
        int col = 0;
        while (getline(iss, token, (char)200))
        {
            node *currentCell = excel.getCell(col, row);
            if (currentCell != nullptr)
            {
                currentCell->data->setData(token);
            }
            col++;
        }
        row++;
    }
    inFile.close();
}
void menu()
{
    system("cls");
    cout << "Enter k to extend row below" << endl;
    cout << "Enter w to extend row above" << endl;
    cout << "Enter b to extend column right" << endl;
    cout << "Enter e to extend colunm left" << endl;
    cout << "Enter d to delete column" << endl;
    cout << "Enter z to delete row" << endl;
    cout << "Enter r to clear column" << endl;
    cout << "Enter c to clear row" << endl;
    cout << "Enter l to insert right shift" << endl;
    cout << "Enter p to insert down shift" << endl;
    cout << "Enter o to delete left shift" << endl;
    cout << "Enter i to delete up shift" << endl;
    cout << "Enter u to calculate range sum" << endl;
    cout << "Enter y to calculate range average" << endl;
    cout << "Enter t to copy" << endl;
    cout << "Enter m to cut" << endl;
    cout << "Enter n to paste" << endl;
    cout << "Enter . to load data" << endl;
    cout << "Enter = for sub functions" << endl;
    cout << "Enter ; yo save data in file" << endl;
    cout << "Enter q to break" << endl;
    cout << "Enter any key to return to main menu";
    getch();
    main();
}
int main()
{
    system("cls");
    int choice;
    cout << "1-Enter Excel" << endl;
    cout << "2-View Keys" << endl;
    cout << "Enter your choice: ";
    cin >> choice;
    if (choice == 1)
    {
        loadData();
        Excel excel(rowCount, columnCount);
        excel.printSheet();
        while (true)
        {
            char key = _getch();
            if (key == 72)
            {
                excel.moveUp();
            }
            if (key == 80)
            {
                excel.moveDown();
            }
            if (key == 75)
            {
                excel.moveLeft();
            }
            if (key == 77)
            {
                excel.moveRight();
            }
            if (key == 32)
            {
                excel.enterData();
            }
            if (key == 'k')
            {
                excel.extendRowBelow();
                excel.printSheet();
            }
            if (key == 'b')
            {
                excel.extendColumnRight();
                excel.printSheet();
            }
            if (key == 'q')
            {
                break;
            }
            if (key == 'd')
            {
                excel.deleteColumn();
                excel.printSheet();
            }
            if (key == 'z')
            {
                excel.deleteRow();
                excel.printSheet();
            }
            if (key == 'c')
            {
                excel.clearRow();
                excel.printSheet();
            }
            if (key == 'r')
            {
                excel.clearColumn();
                excel.printSheet();
            }
            if (key == 'w')
            {
                excel.extendRowUp();
                excel.printSheet();
            }
            if (key == 'e')
            {
                excel.extendColumnLeft();
                excel.printSheet();
            }
            if (key == 'l')
            {
                excel.InsertByRightShift();
                excel.printSheet();
            }
            if (key == 'p')
            {
                excel.InsertByDownShift();
                excel.printSheet();
            }
            if (key == 'o')
            {
                excel.DeleteByLeftShift();
                excel.printSheet();
            }
            if (key == 'i')
            {
                excel.DeleteByUpShift();
                excel.printSheet();
            }
            if (key == 'u')
            {
                int x1, x2, y1, y2;
                cout << endl;
                cout << "Enter x-Coordinte of stating cell: ";
                cin >> x1;
                cout << "Enter y-Coordinte of stating cell: ";
                cin >> y1;
                cout << "Enter x-Coordinte of ending cell: ";
                cin >> x2;
                cout << "Enter y-Coordinte of ending cell: ";
                cin >> y2;
                if (x1 == x2 || y1 == y2)
                {
                    int val = excel.calculateRangeSum(excel.getCell(x1, y1), excel.getCell(x2, y2));
                    cout << "Sum of Range is: " << val << endl;
                    getch();
                    excel.printSheet();
                }
                else
                {
                    cout << "Please enter cells that are either in same row or same column" << endl;
                    end();
                }
            }
            if (key == 'y')
            {
                int x1, x2, y1, y2;
                cout << endl;
                cout << "Enter x-Coordinte of stating cell: ";
                cin >> x1;
                cout << "Enter y-Coordinte of stating cell: ";
                cin >> y1;
                cout << "Enter x-Coordinte of ending cell: ";
                cin >> x2;
                cout << "Enter y-Coordinte of ending cell: ";
                cin >> y2;
                if (x1 == x2 || y1 == y2)
                {
                    float val = excel.calculateRangeAverage(excel.getCell(x1, y1), excel.getCell(x2, y2));
                    cout << "Average of Range is: " << val << endl;
                    excel.printSheet();
                }
                else
                {
                    cout << "Please enter cells that are either in same row or same column" << endl;
                    end();
                }
            }
            if (key == 't')
            {
                int x1, x2, y1, y2;
                cout << endl;
                cout << "Enter x-Coordinte of stating cell: ";
                cin >> x1;
                cout << "Enter y-Coordinte of stating cell: ";
                cin >> y1;
                cout << "Enter x-Coordinte of ending cell: ";
                cin >> x2;
                cout << "Enter y-Coordinte of ending cell: ";
                cin >> y2;
                if (x1 == x2 || y1 == y2)
                {
                    excel.Copy(excel.getCell(x1, y1), excel.getCell(x2, y2));
                    cout << "Your given range has been copied" << endl;
                    excel.printSheet();
                }
                else
                {
                    cout << "Please enter cells that are either in same row or same column" << endl;
                    end();
                }
            }
            if (key == 'm')
            {
                int x1, x2, y1, y2;
                cout << endl;
                cout << "Enter x-Coordinte of stating cell: ";
                cin >> x1;
                cout << "Enter y-Coordinte of stating cell: ";
                cin >> y1;
                cout << "Enter x-Coordinte of ending cell: ";
                cin >> x2;
                cout << "Enter y-Coordinte of ending cell: ";
                cin >> y2;
                if (x1 == x2 || y1 == y2)
                {
                    excel.Cut(excel.getCell(x1, y1), excel.getCell(x2, y2));
                    excel.printSheet();
                }
                else
                {
                    cout << "Please enter cells that are either in same row or same column" << endl;
                    end();
                }
            }
            if (key == 'n')
            {
                string value;
                cout << endl;
                cout << "Enter where you want to paste data i.e in row or column: ";
                cin >> value;
                if (value == "Row" || value == "row" || value == "Column" || value == "column")
                {
                    excel.Paste(value);
                    excel.printSheet();
                }
                else
                {
                    cout << "Please enter correct string" << endl;
                    cout << "Enter any key to contnue";
                    getch();
                    excel.printSheet();
                }
            }
            if (key == ';')
            {
                saveGridToFile("file.txt", excel);
                storeInFile(rowCount, columnCount);
            }
            if (key == '.')
            {
                loadGridFromFile("file.txt", excel);
                excel.printSheet();
            }
            if (key == '=')
            {
                cout << endl;
                char key;
                cout << "Enter s for sum, c for count, a for avergae, m for min and b for max: ";
                cin >> key;
                if (key == 's' || key == 'c' || key == 'a' || key == 'm' || key == 'b')
                {
                    int x1, x2, y1, y2;
                    cout << endl;
                    cout << "Enter x-Coordinte of stating cell: ";
                    cin >> x1;
                    cout << "Enter y-Coordinte of stating cell: ";
                    cin >> y1;
                    cout << "Enter x-Coordinte of ending cell: ";
                    cin >> x2;
                    cout << "Enter y-Coordinte of ending cell: ";
                    cin >> y2;
                    if (x1 == x2 || y1 == y2)
                    {
                        if (key == 's')
                        {
                            excel.calculateSum(excel.getCell(x1, y1), excel.getCell(x2, y2));
                            excel.printSheet();
                        }
                        if (key == 'b')
                        {
                            excel.calculateMax(excel.getCell(x1, y1), excel.getCell(x2, y2));
                            excel.printSheet();
                        }
                        if (key == 'c')
                        {
                            excel.calculateCount(excel.getCell(x1, y1), excel.getCell(x2, y2));
                            excel.printSheet();
                        }
                        if (key == 'm')
                        {
                            excel.calculateMin(excel.getCell(x1, y1), excel.getCell(x2, y2));
                            excel.printSheet();
                        }
                        if (key == 'a')
                        {
                            excel.calculateAverage(excel.getCell(x1, y1), excel.getCell(x2, y2));
                            excel.printSheet();
                        }
                    }
                    else
                    {
                        cout << "Please enter cells that are either in same row or same column" << endl;
                        end();
                    }
                }
                else
                {
                    cout << "Enter correct key" << endl;
                    end();
                }
            }
        }
    }
    else if (choice == 2)
    {
        menu();
    }
    else
    {
        main();
    }
}