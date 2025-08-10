using System;
using System.Runtime.InteropServices;

// Interface COM que será exposta para o VB6
[ComVisible(true)]
[Guid("12345678-1234-1234-1234-123456789012")]
[InterfaceType(ComInterfaceType.InterfaceIsIDispatch)]
public interface IDataExchange
{
    [DispId(1)]
    string GetMessage();

    [DispId(2)]
    void SetMessage(string message);

    [DispId(3)]
    int CalculateSum(int a, int b);

    [DispId(4)]
    string ProcessData(string input);
}

// Implementação do componente COM
[ComVisible(true)]
[Guid("87654321-4321-4321-4321-210987654321")]
[ClassInterface(ClassInterfaceType.None)]
[ProgId("MyApp.DataExchange")]
public class DataExchange : IDataExchange
{
    private string _storedMessage = "Mensagem inicial do C#";

    public string GetMessage()
    {
        return _storedMessage;
    }

    public void SetMessage(string message)
    {
        _storedMessage = message;
        Console.WriteLine($"C# recebeu: {message}");
    }

    public int CalculateSum(int a, int b)
    {
        int result = a + b;
        Console.WriteLine($"C# calculou: {a} + {b} = {result}");
        return result;
    }

    public string ProcessData(string input)
    {
        string processed = $"Processado em C#: {input.ToUpper()} - {DateTime.Now}";
        Console.WriteLine($"C# processou: {input} -> {processed}");
        return processed;
    }
}

// Classe para registrar/desregistrar o componente COM
[ComVisible(false)]
public class ComRegistration
{
    [ComRegisterFunction]
    public static void RegisterFunction(Type t)
    {
        Console.WriteLine($"Registrando componente COM: {t.Name}");
    }

    [ComUnregisterFunction]
    public static void UnregisterFunction(Type t)
    {
        Console.WriteLine($"Desregistrando componente COM: {t.Name}");
    }
}