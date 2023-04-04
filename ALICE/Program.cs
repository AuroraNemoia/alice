using System;
using System.Collections.Generic;
using System.Linq;
using AForge.Video.DirectShow;
using System.Net.NetworkInformation;
using System.Windows.Forms;
using System.IO;
using System.Net;
using System.IO.Compression;
using System.Reflection;
using System.Threading;
using System.Diagnostics;
using System.Text;
using System.Threading.Tasks;

namespace ALICE
{
    internal class Program
    {
        static void Main(string[] args)
        {
            // sanity checks, are pre-requisites available?
            if (CheckIfWebcamExists() == false)
            {
                ShowErrorMessageBox("No webcam found!", "Error");
                Environment.Exit(1);
            }

            if (PingTwitch() == false)
            {
                ShowErrorMessageBox("Cannot reach Twitch.TV!", "Error");
                Environment.Exit(1);
            }
            // download and extract obs portable:
            string obs_url = "https://github.com/obsproject/obs-studio/releases/download/29.0.2/OBS-Studio-29.0.2-Full-x64.zip";
            string vcredist_url = "https://aka.ms/vs/17/release/vc_redist.x64.exe";
            string obs_filename = "OBS-Studio-29.0.2-Full-x64.zip";
            string vc_redist_filename = "vc_redist.x64.exe";
            string obsExecutablePath = Path.Combine("bin", "64bit", "obs64.exe");
            string executablePath = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
            string obs_filePath = Path.Combine(executablePath, obs_filename);
            string vc_redist_filepath = Path.Combine(executablePath, vc_redist_filename);
            string firstWebcamDeviceId = GetFirstWebcamDeviceID();

            if (firstWebcamDeviceId != null)
            {
                Console.WriteLine("First webcam device ID: " + firstWebcamDeviceId);
            }
            else
            {
                Console.WriteLine("No webcam devices found.");
                Environment.Exit(1);
            }

            
            if (!File.Exists(obs_filePath))
            {
                DownloadFile(obs_url, Path.Combine(executablePath, obs_filename));
                ExtractZipFile(Path.Combine(executablePath, obs_filename), executablePath);

            }
            else
            {
                Console.WriteLine("File already exists. Skipping download.");
            }

            Console.WriteLine("Download and extraction complete.");
            DownloadFile(vcredist_url, Path.Combine(executablePath, vc_redist_filename));
            RunVcRedistInstallerSilently(vc_redist_filepath);

            // Start OBS
            Process obsProcess = StartOBS(obsExecutablePath);

            // Wait for 5 seconds
            Thread.Sleep(5000);

            // Kill OBS process
            KillProcess(obsProcess);

            Console.WriteLine("OBS process killed.");

        }

        static bool PingTwitch()
        {
            try
            {
                Ping pingSender = new Ping();
                PingReply reply = pingSender.Send("twitch.tv");

                return reply.Status == IPStatus.Success;
            }
            catch (PingException)
            {
                return false;
            }
        }

        static bool CheckIfWebcamExists()
        {
            // Enumerate video devices
            FilterInfoCollection videoDevices = new FilterInfoCollection(FilterCategory.VideoInputDevice);

            // Check if any devices are found
            if (videoDevices.Count > 0)
            {
                return true;
            }
            else
            {
                return false;
            }
        }
        static void ShowErrorMessageBox(string message, string title)
        {
            MessageBox.Show(
                message,
                title,
                MessageBoxButtons.OK,
                MessageBoxIcon.Error
            );
        }
        static void DownloadFile(string obs_url, string obs_filePath)
        {
            using (WebClient client = new WebClient())
            {
                Console.WriteLine($"Downloading file from {obs_url}...");
                client.DownloadFile(obs_url, obs_filePath);
                Console.WriteLine("Download complete.");
            }
        }

        static void ExtractZipFile(string zipFilePath, string destinationFolderPath)
        {
            Console.WriteLine($"Extracting file to {destinationFolderPath}...");
            ZipFile.ExtractToDirectory(zipFilePath, destinationFolderPath);
            Console.WriteLine("Extraction complete.");
        }

        static Process StartOBS(string obsExecutablePath)
        {
            string obsWorkingDirectory = Path.GetDirectoryName(obsExecutablePath);
            
            ProcessStartInfo startInfo = new ProcessStartInfo
            {
                FileName = obsExecutablePath,
                Arguments = "--portable --minimize-to-tray",
                WorkingDirectory = obsWorkingDirectory,
                UseShellExecute = false
            };

            Process obsProcess = new Process
            {
                StartInfo = startInfo
            };

            obsProcess.Start();
            Console.WriteLine("OBS started.");

            return obsProcess;
        }

        static void KillProcess(Process process)
        {
            if (process != null && !process.HasExited)
            {
                process.Kill();
                process.WaitForExit();
            }
        }

        static void RunVcRedistInstallerSilently(string installerPath)
        {
            ProcessStartInfo startInfo = new ProcessStartInfo
            {
                FileName = installerPath,
                Arguments = "/install /quiet /norestart",
                UseShellExecute = false,
                CreateNoWindow = true
            };

            Process installerProcess = new Process
            {
                StartInfo = startInfo
            };

            Console.WriteLine("Running vc_redist.x64.exe installer silently...");
            installerProcess.Start();
            installerProcess.WaitForExit();
            Console.WriteLine("Installer finished.");
        }

        static string GetFirstWebcamDeviceID()
        {
            FilterInfoCollection videoDevices = new FilterInfoCollection(FilterCategory.VideoInputDevice);

            if (videoDevices.Count > 0)
            {
                FilterInfo firstDevice = videoDevices[0];
                return firstDevice.MonikerString;
            }
            else
            {
                return null;
            }
        }
    }
}
