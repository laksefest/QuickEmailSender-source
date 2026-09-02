using System;
using System.Runtime.InteropServices;
using System.Threading;
using System.Threading.Tasks;

namespace SapOrderReader;

internal static class KeyboardNavigation
{
    private const uint KeyEventKeyUp = 0x0002;

    private const byte VkTab = 0x09;
    private const byte VkEnter = 0x0D;

    [DllImport("user32.dll", SetLastError = true)]
    private static extern void keybd_event(
        byte virtualKey,
        byte scanCode,
        uint flags,
        UIntPtr extraInfo);

    public static async Task TabEnterAsync(
        int tabCount,
        int delayMilliseconds = 140)
    {
        if (tabCount < 0)
        {
            throw new ArgumentOutOfRangeException(
                nameof(tabCount),
                "Tab count cannot be negative.");
        }

        for (int i = 0; i < tabCount; i++)
        {
            PressKey(VkTab);
            await Task.Delay(delayMilliseconds);
        }

        PressKey(VkEnter);
        await Task.Delay(delayMilliseconds);
    }

    private static void PressKey(byte virtualKey)
    {
        keybd_event(
            virtualKey,
            0,
            0,
            UIntPtr.Zero);

        Thread.Sleep(80);

        keybd_event(
            virtualKey,
            0,
            KeyEventKeyUp,
            UIntPtr.Zero);

        Thread.Sleep(80);
    }
}
