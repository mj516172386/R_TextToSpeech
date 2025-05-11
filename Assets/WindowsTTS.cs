using System;
using System.Diagnostics;
using System.Linq;
using UnityEngine;
using UnityEngine.UI;
using Debug = UnityEngine.Debug;

[Serializable]
public enum VoiceType
{
    ChineseFemale,
    EnglishFemale,
    EnglishMale
}

public class WindowsTTS : MonoBehaviour
{
    [Header("UI Elements")]
    public InputField mInput;
    public Button startButton;
    public Button stopButton;
    public Dropdown voiceDropdown;

    [Header("Settings")]
    [Tooltip("Default voice selection")]
    public VoiceType defaultVoice = VoiceType.ChineseFemale;

    private Process ttsProcess;
    private bool isPlaying = false;
    private VoiceType currentVoice;

    void Start()
    {
        InitializeUI();
        ListInstalledVoices();
        currentVoice = defaultVoice;
    }

    void InitializeUI()
    {
        /*
         Microsoft Huihui Desktop | Gender: Female
Microsoft Zira Desktop | Gender: Female
Microsoft David Desktop | Gender: Male

         */
        startButton.onClick.AddListener(StartConversion);
        stopButton.onClick.AddListener(StopTTS);

        // Setup voice dropdown
     
    }

    void OnVoiceSelected(int index)
    {
        currentVoice = (VoiceType)index;
    }

    void ListInstalledVoices()
    {
        try
        {
            string command = @"
                Add-Type -AssemblyName System.Speech;
                $synth = New-Object System.Speech.Synthesis.SpeechSynthesizer;
                $synth.GetInstalledVoices() | ForEach-Object {
                    Write-Output ($_.VoiceInfo.Name + ' | Gender: ' + $_.VoiceInfo.Gender)
                }";

            ExecutePowerShellCommand(command, output => {
                UnityEngine.Debug.Log("=== 可用语音列表 ===\n" + output);
                string[] voiceEntries = output.Split('\n').Select(entry => entry.Trim()).ToArray();
                voiceDropdown.ClearOptions();
                voiceDropdown.AddOptions(voiceEntries.ToList());
                voiceDropdown.onValueChanged.AddListener(OnVoiceSelected);
            });
        }
        catch (Exception e)
        {
            Debug.LogError($"获取语音列表失败: {e.Message}");
        }
    }

    public void StartConversion()
    {
        if (string.IsNullOrEmpty(mInput.text))
        {
            Debug.LogWarning("输入文本为空");
            return;
        }

        if (isPlaying)
        {
            StopTTS();
        }
        StartTTS(mInput.text);
    }
  
    void StartTTS(string text)
    {
        CleanupProcess();

        try
        {
            string voiceName = GetVoiceName(currentVoice);
            string command = $@"
                Add-Type -AssemblyName System.Speech;
                $speak = New-Object System.Speech.Synthesis.SpeechSynthesizer;
                $speak.SelectVoice('{voiceName}');
                $speak.Volume = 100;
                $speak.Rate = 0;
                $speak.Speak('{EscapeText(text)}')";

            ExecutePowerShellCommand(command, null, true);
            isPlaying = true;
        }
        catch (Exception e)
        {
            Debug.LogError($"TTS启动失败: {e.Message}");
            CleanupProcess();
        }
    }

    void ExecutePowerShellCommand(string command, Action<string> callback = null, bool isTTS = false)
    {
        ProcessStartInfo startInfo = new ProcessStartInfo()
        {
            FileName = "powershell.exe",
            Arguments = command,
            UseShellExecute = false,
            CreateNoWindow = true,
            RedirectStandardOutput = true,
            RedirectStandardError = true
        };

        Process process = new Process { StartInfo = startInfo };

        if (isTTS)
        {
            ttsProcess = process;
            ttsProcess.EnableRaisingEvents = true;
            ttsProcess.Exited += (sender, args) => isPlaying = false;
        }

        process.Start();

        if (callback != null)
        {
            string output = process.StandardOutput.ReadToEnd();
            process.WaitForExit();
            callback(output);
        }
    }

    string GetVoiceName(VoiceType voiceType)
    {
        return voiceType switch
        {
            VoiceType.ChineseFemale => "Microsoft Huihui Desktop",
            VoiceType.EnglishFemale => "Microsoft Zira Desktop",
            VoiceType.EnglishMale => "Microsoft David Desktop",
            _ => "Microsoft Huihui Desktop"
        };
    }

    void StopTTS()
    {
        if (!isPlaying) return;
        CleanupProcess();
        Debug.Log("TTS已停止");
    }

    void CleanupProcess()
    {
        if (ttsProcess != null)
        {
            try
            {
                if (!ttsProcess.HasExited)
                {
                    ttsProcess.Kill();
                }
                ttsProcess.Dispose();
            }
            catch (Exception e)
            {
                Debug.LogWarning($"进程清理异常: {e.Message}");
            }
            finally
            {
                ttsProcess = null;
                isPlaying = false;
            }
        }
    }

    string EscapeText(string text)
    {
        return text.Replace("'", "''")
                  .Replace("$", "`$")
                  .Replace("\"", "`\"")
                  .Replace("\r\n", " ")
                  .Replace("\n", " ");
    }

    void OnDestroy()
    {
        CleanupProcess();
    }
}