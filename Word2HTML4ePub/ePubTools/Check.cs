﻿using System.Text;
using System.Windows.Forms;

namespace Word2HTML4ePub
{
    public partial class ePubTools
    {

        /// <summary>
        /// Automation de ePubCheck en local1
        /// </summary>
        /// <param name="filename">l'ePub à vérifier</param>
        /// <returns>null si OK, sinon les erreurs</returns>
        public static string CheckEPub(string filename)
        {
            StringBuilder generalstring = new StringBuilder();
            StringBuilder output = new StringBuilder();
            System.Diagnostics.Process proc = new System.Diagnostics.Process();// processus de verification
            proc.StartInfo = new System.Diagnostics.ProcessStartInfo("\"" + JavaPath + "\"", "-jar " + "\"" + EpubCheckPath + "\" \"" + filename + "\"");
            proc.StartInfo.CreateNoWindow = true; // ne pas afficher de fenetre
            proc.StartInfo.UseShellExecute = false; // On désactive le shell
            proc.StartInfo.RedirectStandardOutput = true; // On redirige la sortie standard
            proc.StartInfo.RedirectStandardError = true;
            proc.Start(); // Démarrage du processus

            while (true)
            {
                output.Append(proc.StandardError.ReadLine());
                if (proc.HasExited)
                    break;
            }

            generalstring.Append(proc.StandardOutput.ReadLine());
            proc.WaitForExit(); // Attente de la fin de la commande

            proc.Close(); // Libération des ressources
            if (output.Length == 0)
                return null;
            else
                return generalstring + "\r\n" + output.ToString();
        }

    }
}