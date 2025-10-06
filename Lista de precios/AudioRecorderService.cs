using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Lista_de_precios
{
    internal class AudioRecorderService
    {

        private bool isRecording = false;
        private string audioFilePath;

        public async Task<bool> StartRecordingAsync()
        {
            try
            {
                // Solicitar permisos
                var status = await Permissions.RequestAsync<Permissions.Microphone>();
                var storageStatus = await Permissions.RequestAsync<Permissions.StorageWrite>();

                if (status != PermissionStatus.Granted || storageStatus != PermissionStatus.Granted)
                {
                    return false;
                }

                // Crear ruta del archivo
                audioFilePath = Path.Combine(FileSystem.CacheDirectory, $"audio_{DateTime.Now:yyyyMMdd_HHmmss}.3gp");

#if ANDROID
                var activity = Platform.CurrentActivity;
                var recorder = new Android.Media.MediaRecorder();
                
                recorder.SetAudioSource(Android.Media.AudioSource.Mic);
                recorder.SetOutputFormat(Android.Media.OutputFormat.ThreeGpp);
                recorder.SetAudioEncoder(Android.Media.AudioEncoder.AmrNb);
                recorder.SetOutputFile(audioFilePath);
                
                recorder.Prepare();
                recorder.Start();
                
                isRecording = true;
                
                // Guardar referencia al recorder
                activity.GetType().GetField("_mediaRecorder", 
                    System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Instance)?
                    .SetValue(activity, recorder);
#endif
                return true;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error al iniciar grabación: {ex.Message}");
                return false;
            }
        }

        public async Task<string> StopRecordingAsync()
        {
            if (!isRecording) return null;

            try
            {
#if ANDROID
                var activity = Platform.CurrentActivity;
                var recorder = activity.GetType().GetField("_mediaRecorder",
                    System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Instance)?
                    .GetValue(activity) as Android.Media.MediaRecorder;
                
                if (recorder != null)
                {
                    recorder.Stop();
                    recorder.Release();
                    recorder = null;
                }
#endif
                isRecording = false;

                // Verificar si el archivo se creó
                if (File.Exists(audioFilePath))
                {
                    return audioFilePath;
                }

                return null;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error al detener grabación: {ex.Message}");
                return null;
            }
        }

        public bool IsRecording => isRecording;
    }
}
