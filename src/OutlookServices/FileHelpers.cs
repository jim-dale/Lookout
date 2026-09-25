namespace OutlookServices;

using System;
using System.IO;

/// <summary>
/// Provides helper methods for working with files and file paths.
/// </summary>
public static class FileHelpers
{
    /// <summary>
    /// Expands environment variables in a path and returns its absolute path.
    /// </summary>
    /// <param name="s">The path that may contain environment variables.</param>
    /// <returns>The expanded absolute path, or an empty string if <paramref name="s"/> is null, empty, or whitespace.</returns>
    public static string GetPathWithEnvVars(string s)
    {
        string result = string.Empty;

        if (string.IsNullOrWhiteSpace(s) == false)
        {
            result = Environment.ExpandEnvironmentVariables(s);
            result = Path.GetFullPath(result);
        }

        return result;
    }

    /// <summary>
    /// Trims a path and replaces invalid file-name characters with the specified character.
    /// </summary>
    /// <param name="path">The file name to clean.</param>
    /// <param name="replaceChar">The character used to replace invalid file-name characters.</param>
    /// <returns>The cleaned file name.</returns>
    /// <exception cref="ArgumentException">Thrown when <paramref name="path"/> is null, empty, or whitespace.</exception>
    public static string CleanFileName(string path, char replaceChar = '_')
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(path);

        string result = string.Empty;

        if (string.IsNullOrWhiteSpace(path) == false)
        {
            result = path.Trim();

            foreach (char c in Path.GetInvalidFileNameChars())
            {
                result = result.Replace(c, replaceChar);
            }
        }

        return result;
    }

    /// <summary>
    /// Gets a file name that does not conflict with an existing file in the specified folder.
    /// </summary>
    /// <param name="folder">The folder in which the file will be created.</param>
    /// <param name="fileName">The requested file name.</param>
    /// <returns>A unique file path in <paramref name="folder"/>.</returns>
    public static string GetUniqueFilePath(DirectoryInfo folder, string fileName)
    {
        return GetUniqueFilePath(folder.FullName, fileName);
    }

    /// <summary>
    /// Gets a file name that does not conflict with an existing file in the specified folder.
    /// </summary>
    /// <param name="folder">The path to the folder in which the file will be created.</param>
    /// <param name="fileName">The requested file name.</param>
    /// <returns>A unique file path in <paramref name="folder"/>.</returns>
    public static string GetUniqueFilePath(string folder, string fileName)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(fileName);

        // Get just the file name and extension
        string baseFileName = Path.GetFileNameWithoutExtension(fileName);
        string extension = Path.GetExtension(fileName);

        fileName = baseFileName + extension;
        string result = Path.Combine(folder, fileName);

        int counter = 1;
        while (File.Exists(result))
        {
            fileName = $"{baseFileName} ({counter}){extension}";
            result = Path.Combine(folder, fileName);
            ++counter;
        }

        return result;
    }
}
