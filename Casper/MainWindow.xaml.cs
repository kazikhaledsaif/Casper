using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.Speech.Recognition;
using System.Speech.Synthesis;

namespace Casper
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private SpeechRecognitionEngine speechRecognitionEngine = new SpeechRecognitionEngine(new System.Globalization.CultureInfo("en-US"));
        private SpeechSynthesizer speechSynthesizer = new SpeechSynthesizer();
        private Boolean wake = false;

        public MainWindow()
        {
            InitializeComponent();


            try
            {
                
                //hook to events
               speechRecognitionEngine.AudioLevelUpdated += new EventHandler<AudioLevelUpdatedEventArgs>(engine_AudioLavelUpdated);
               speechRecognitionEngine.SpeechRecognized +=new EventHandler<SpeechRecognizedEventArgs>(engine_SpeechRecognized);


                
                //load dictionary
                LoadGrammerAndCommands();

                //using system default microphone
                speechRecognitionEngine.SetInputToDefaultAudioDevice();

                speechRecognitionEngine.InitialSilenceTimeout = TimeSpan.FromSeconds(3);
                speechRecognitionEngine.BabbleTimeout = TimeSpan.FromSeconds(2);
                speechRecognitionEngine.EndSilenceTimeout = TimeSpan.FromSeconds(1);
                speechRecognitionEngine.EndSilenceTimeoutAmbiguous = TimeSpan.FromSeconds(1.5);
              
                //start listening
                speechRecognitionEngine.RecognizeAsync(RecognizeMode.Multiple);
                speechSynthesizer.SpeakCompleted += new EventHandler<SpeakCompletedEventArgs>(casper_SpeakCompleted);

                if (speechSynthesizer.State == SynthesizerState.Speaking)
                {
                    speechSynthesizer.SpeakAsyncCancelAll();
                }
               


            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message,"Voice Recongnation Failed");
            }

        }

        private void casper_SpeakCompleted(object sender, SpeakCompletedEventArgs e)
        {
            if (speechSynthesizer.State == SynthesizerState.Speaking)
            {
                speechSynthesizer.SpeakAsyncCancelAll();
            }

        }

        private void engine_SpeechRecognized(object sender, SpeechRecognizedEventArgs e)
        {
            string speech = e.Result.Text;
            casperTB.AppendText(speech + "\n");
        

                switch (speech)
                {
                    case "hello":
                        speechSynthesizer.SpeakAsync("Hello Sir");
                        break;
                    case "what time it is":
                        speechSynthesizer.SpeakAsync(DateTime.Now.ToString("hh:mm tt zz"));
                        break;
                    case "what is today":
                        speechSynthesizer.SpeakAsync(DateTime.Now.ToString("dd MMMM yyyy  "));
                        break;
                    case "how are you":
                        speechSynthesizer.SpeakAsync("i am fine, Sir");
                        break;
                    case "open google":
                        System.Diagnostics.Process.Start("http://www.google.com");
                        break;
                    case "open facebook":
                        System.Diagnostics.Process.Start("http://www.facebook.com");
                        break;
                    case "open office":
                        System.Diagnostics.Process.Start(@"C:\Program Files\Microsoft Office\Office16/WINWORD.EXE");
                        break;

                }


    



        }

        private void engine_AudioLavelUpdated(object sender, AudioLevelUpdatedEventArgs e)
        {
            casperPB.Value = e.AudioLevel;
        }



        private void LoadGrammerAndCommands()
        {

            try
            {
                Choices choices = new Choices();
                string[] lines = File.ReadAllLines(Environment.CurrentDirectory + "\\Commands.txt");
                choices.Add(lines);
                Grammar grammar = new Grammar(new GrammarBuilder(choices));
                speechRecognitionEngine.LoadGrammar(grammar);
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message);
               
            }
        }



    }
}
