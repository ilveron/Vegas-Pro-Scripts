using System;
using System.IO;
using System.Windows.Forms;
using ScriptPortal.Vegas;

public class EntryPoint
{
    public void FromVegas(Vegas vegas)
    {
        // === Choose renderer ===
        Renderer renderer = PromptRendererChoice(vegas);
        if (renderer == null)
        {
            MessageBox.Show("No renderer selected. Export cancelled.", "Export Audio Tracks");
            return;
        }

        // === Choose template ===
        RenderTemplate template = PromptTemplateChoice(renderer);
        if (template == null)
        {
            MessageBox.Show("No template selected. Export cancelled.", "Export Audio Tracks");
            return;
        }

        // === Determine output folder ===
        string projectPath = vegas.Project.FilePath;
        string outputFolder = Path.Combine(
            string.IsNullOrEmpty(projectPath)
                ? Environment.GetFolderPath(Environment.SpecialFolder.Desktop)
                : Path.GetDirectoryName(projectPath),
            "AudioExports"
        );
        Directory.CreateDirectory(outputFolder);

        int trackCounter = 1;

        // === Process each audio track ===
        foreach (Track track in vegas.Project.Tracks)
        {
            if (!(track is AudioTrack))
                continue;

            AudioTrack audioTrack = (AudioTrack)track;

            // Mute all audio tracks
            foreach (Track t in vegas.Project.Tracks)
            {
                if (t is AudioTrack)
                    t.Mute = true;
            }

            // Unmute current track
            audioTrack.Mute = false;

            string trackName = string.IsNullOrEmpty(audioTrack.Name)
                ? "Track_" + trackCounter
                : audioTrack.Name;

            string safeName = MakeSafeFilename(trackName);
            string outputPath = Path.Combine(outputFolder, safeName + ".wav");

            // Render
            RenderArgs args = new RenderArgs
            {
                OutputFile = outputPath,
                RenderTemplate = template
            };
            vegas.Render(args);

            // Mute again
            audioTrack.Mute = true;
            trackCounter++;
        }

        MessageBox.Show(
            "Export completed!\nFiles saved in:\n" + outputFolder,
            "Export Audio Tracks"
        );
    }

    // --- Utilities ---

    private string MakeSafeFilename(string name)
    {
        if (string.IsNullOrEmpty(name)) return "UnnamedTrack";
        foreach (char c in Path.GetInvalidFileNameChars())
            name = name.Replace(c, '_');
        return name.Trim();
    }

    private Renderer PromptRendererChoice(Vegas vegas)
    {
        using (Form form = new Form())
        {
            form.Text = "Select Audio Renderer";
            form.Width = 400;
            form.Height = 150;
            form.FormBorderStyle = FormBorderStyle.FixedDialog;
            form.StartPosition = FormStartPosition.CenterScreen;
            form.MinimizeBox = false;
            form.MaximizeBox = false;

            Label label = new Label() { Left = 10, Top = 10, Text = "Available audio renderers:", Width = 350 };
            ComboBox combo = new ComboBox() { Left = 10, Top = 35, Width = 360, DropDownStyle = ComboBoxStyle.DropDownList };

            foreach (Renderer r in vegas.Renderers)
            {
                if (r.FileTypeName.ToLower().Contains("wav") ||
                    r.FileTypeName.ToLower().Contains("wave") ||
                    r.FileTypeName.ToLower().Contains("audio"))
                {
                    combo.Items.Add(r.FileTypeName);
                }
            }

            if (combo.Items.Count == 0)
            {
                MessageBox.Show("No audio renderers found!", "Export Audio Tracks");
                return null;
            }

            combo.SelectedIndex = 0;
            Button okButton = new Button() { Text = "OK", Left = 210, Width = 75, Top = 70, DialogResult = DialogResult.OK };
            Button cancelButton = new Button() { Text = "Cancel", Left = 295, Width = 75, Top = 70, DialogResult = DialogResult.Cancel };

            form.Controls.Add(label);
            form.Controls.Add(combo);
            form.Controls.Add(okButton);
            form.Controls.Add(cancelButton);
            form.AcceptButton = okButton;
            form.CancelButton = cancelButton;

            if (form.ShowDialog() == DialogResult.OK)
            {
                string selected = combo.SelectedItem.ToString();
                foreach (Renderer r in vegas.Renderers)
                {
                    if (r.FileTypeName == selected)
                        return r;
                }
            }
            return null;
        }
    }

    private RenderTemplate PromptTemplateChoice(Renderer renderer)
    {
        using (Form form = new Form())
        {
            form.Text = "Select Render Template";
            form.Width = 500;
            form.Height = 200;
            form.FormBorderStyle = FormBorderStyle.FixedDialog;
            form.StartPosition = FormStartPosition.CenterScreen;
            form.MinimizeBox = false;
            form.MaximizeBox = false;

            Label label = new Label() { Left = 10, Top = 10, Text = "Available templates for " + renderer.FileTypeName + ":", Width = 460 };
            ComboBox combo = new ComboBox() { Left = 10, Top = 35, Width = 460, DropDownStyle = ComboBoxStyle.DropDownList };

            foreach (RenderTemplate t in renderer.Templates)
            {
                if (t.IsValid())
                    combo.Items.Add(t.Name);
            }

            if (combo.Items.Count == 0)
            {
                MessageBox.Show("No valid templates found for this renderer!", "Export Audio Tracks");
                return null;
            }

            combo.SelectedIndex = 0;
            Button okButton = new Button() { Text = "OK", Left = 300, Width = 80, Top = 100, DialogResult = DialogResult.OK };
            Button cancelButton = new Button() { Text = "Cancel", Left = 390, Width = 80, Top = 100, DialogResult = DialogResult.Cancel };

            form.Controls.Add(label);
            form.Controls.Add(combo);
            form.Controls.Add(okButton);
            form.Controls.Add(cancelButton);
            form.AcceptButton = okButton;
            form.CancelButton = cancelButton;

            if (form.ShowDialog() == DialogResult.OK)
            {
                string selected = combo.SelectedItem.ToString();
                foreach (RenderTemplate t in renderer.Templates)
                {
                    if (t.Name == selected)
                        return t;
                }
            }
            return null;
        }
    }
}
