using ScriptPortal.Vegas;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Text.Json;
using System.Threading.Tasks;
using System.Windows.Forms;
using static CTAImport.CTAIMainForm;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.TaskbarClock;

namespace CTAImport
{
    public partial class CTAIMainForm : UserControl
    {
       Vegas myVegas;
        string CTAVersion = "V01.017";

        bool debug = false;

        //public Vegas myVegas;

        private static String RGBConverter(System.Drawing.Color c)
        {
            return "RGB(" + c.R.ToString() + "," + c.G.ToString() + "," + c.B.ToString() + ")";
        }

        public CTAIMainForm(Vegas vegas)
        {

            myVegas = vegas;
            SaveLogFile(MethodBase.GetCurrentMethod(), "Started");

#if DEBUG
            this.debug = true;
#endif

            InitializeComponent();
            lblMain1.Text = "Import CartoonAnimator projects into VEGAS";
            lblMain2.Text = "In order to import an animated project into VEGAS, kindly import the ";
            lblMain3.Text = "Cartoon Animator JSON file created by the <Export to After Effects> command";

            lblParam01.Text = "Layer Offset x";
            lblParam02.Text = "Layer Offset y";
            lblParam03.Text = "Layer Offset z";
            lblParam04.Text = "Cam Offset x";
            lblParam05.Text = "Cam Offset y";
            lblParam06.Text = "Cam Offset z";
            lblParam07.Text = "Layer Factor x";
            lblParam08.Text = "Layer Factor y";
            lblParam09.Text = "Layer Factor z";
            lblParam10.Text = "Cam Factor x";
            lblParam11.Text = "Cam Factor y";
            lblParam12.Text = "Cam Factor z";
            lblParam13.Text = "CamRatio Factor";
            if (this.debug)
            {
                lbl_copyright.Text = "© Harold Linke 2023 - Version " + this.CTAVersion + "(DEBUG)";
            }
            else
            {
                lbl_copyright.Text = "© Harold Linke 2023 - Version " + this.CTAVersion;
            }
            groupBox1.Text = "Finetune Keyframes";
            numParam01.Value = 0;
            numParam02.Value = 0;
            numParam03.Value = 0;
            numParam04.Value = 0;
            numParam05.Value = 0;
            numParam06.Value = 0;
            numParam07.Value = 100;
            numParam08.Value = 100;
            numParam09.Value = 100;
            numParam10.Value = 100;
            numParam11.Value = 100;
            numParam12.Value = 100;
            numParam13.Value = 100;

            btnImportCTA.Text = "Import CTA Project";
            btnRecalcCameraKeyFrames.Text = "Recalc Camera KeyFrames";
            btnRecalcCameraKeyFrames.Visible = false;
            setcolors();
        }

        public void setcolors()
        {
            this.BackColor = Color.Empty; // this.lblMain1.BackColor;
            //ShowLogMessage(RGBConverter(this.lblMain1.BackColor));
            groupBox1.ForeColor = this.lblMain1.ForeColor;
        }

        public void btnImportCTA_Click(object sender, EventArgs e)
        {
            string s = "CTA Import - Version:" + this.CTAVersion;
            SaveLogFile(MethodBase.GetCurrentMethod(), s, true);
            using (UndoBlock undo = new UndoBlock("Import CTA Project"))
            {
                ImportCTAProject(myVegas);
            }
        }

        public void btnRecalcCameraKeyFrames_Click(object sender, EventArgs e)
        {
            try
            {
                using (UndoBlock undo = new UndoBlock("Recalc Camera KeyFrames"))
                {
                    SaveLogFile(MethodBase.GetCurrentMethod(), "Started");
                    Camera_Create(CameraVidtrack, myCreate.camera, 0);
                    int trackstartindex = CameraVidtrack.Index;
                    trackstartindex += 1; // index of track with first layer

                    foreach (Comp jsonComp in myCreate.comps)
                    {
                        string s = "------------------------------------------------------------------------------------------------" + "\r\n";
                        s += "* Update Layer started: " + jsonComp.name + "\r\n";
                        s += "* HasAudio:" + jsonComp.hasAudio + "\r\n";
                        s += "* HasVideo:" + jsonComp.hasVideo + "\r\n";
                        s += "* trackstartindex:" + trackstartindex + "\r\n";
                        s += "------------------------------------------------------------------------------------------------" + "\r\n";
                        SaveLogFile(MethodBase.GetCurrentMethod(), s);
                        ShowLogMessage("Update Track: " + jsonComp.name + "\r\n");

                        Layer_Create(jsonComp, ref trackstartindex, updateonly: true);
                    }
                }
            }
            catch (Exception exc)
            {
                ShowLogMessage("**Update-ERROR** " + exc.Message + "\r\n");
            }
        }


        // Root myDeserializedClass = JsonConvert.DeserializeObject<Root>(myJsonResponse);
        public class Background
        {
            public int count { get; set; }
            public string filename { get; set; }
            public bool hasAudio { get; set; }
            public bool hasVideo { get; set; }
        }

        public class Camera
        {
            public int camLens { get; set; }
            public List<PosObj> posObjs { get; set; }
            public List<RotObj> rotObjs { get; set; }
        }

        public class Comp
        {
            public CompObj compObj { get; set; }
            public CompScale compScale { get; set; }
            public bool hasAudio { get; set; }
            public bool hasVideo { get; set; }
            public int id { get; set; }
            public bool isThreeD { get; set; }
            public List<object> layerObjs { get; set; }
            public string name { get; set; }
            public string replaceTgt { get; set; }
        }

        public class CompObj
        {
            public int count { get; set; }
            public List<object> depth { get; set; }
            public string fileName { get; set; }
            public List<Pivot> pivot { get; set; }
            public List<PosObj> posObjs { get; set; }
            public List<RotObj> rotObjs { get; set; }
            public List<ScaleObj> scaleObjs { get; set; }
        }

        public class CompScale
        {
            public double scaleX { get; set; }
            public int scaleY { get; set; }
            public double scaleZ { get; set; }
        }

        public class Create
        {
            public Background background { get; set; }
            public Camera camera { get; set; }
            public List<Comp> comps { get; set; }
            public int endFrame { get; set; }
            public double sceneRatio { get; set; }
            public int startFrame { get; set; }
        }

        public class Pivot
        {
            public int time { get; set; }
            public double x { get; set; }
            public double y { get; set; }
            public double z { get; set; }
        }

        public class PosObj
        {
            public int time { get; set; }
            public double x { get; set; }
            public double y { get; set; }
            public double z { get; set; }
        }

        public class ProjectAudio
        {
            public bool music { get; set; }
            public bool soundFX1 { get; set; }
            public bool soundFX2 { get; set; }
        }

        public class Root
        {
            public Create Create { get; set; }
            public int fps { get; set; }
            public int height { get; set; }
            public bool loadBgm { get; set; }
            public string name { get; set; }
            public bool preMulti { get; set; }
            public ProjectAudio projectAudio { get; set; }
            public bool smartType { get; set; }
            public string version { get; set; }
            public ViewCornerLT viewCornerLT { get; set; }
            public ViewCornerRB viewCornerRB { get; set; }
            public int width { get; set; }
        }

        public class RotObj
        {
            public int time { get; set; }
            public int value { get; set; }
            public double x { get; set; }
            public double y { get; set; }
            public double z { get; set; }
        }

        public class ScaleObj
        {
            public int time { get; set; }
            public double x { get; set; }
            public double y { get; set; }
            public double z { get; set; }
        }

        public class ViewCornerLT
        {
            public double x { get; set; }
            public double y { get; set; }
            public double z { get; set; }
        }

        public class ViewCornerRB
        {
            public double x { get; set; }
            public double y { get; set; }
            public double z { get; set; }
        }

        public class SceneRatio
        {
            public double scene;
            public double width;
            public double height;
        }

        public class AECamera
        {
            public double focus;
            public string name = "AECamera";
            public PosObj pos;
            public double zoom;
            public double focusDistance;
        }


        int orgWidth;
        int orgHeight;
        double timeShift;
        public double cameraFLM = 53.33333334;
        public double standardZinCTACam = -2099.2021484375;
        public Root CTAScene;
        string jsonString;
        string FileName = "";

        string maindir;

        public Create myCreate;
        double fps;
        string SceneName;
        int SceneWidth;
        int SceneHeight;
        double scaletime;
        //double scalex;
        //double offsetx;
        //double offsety;
        //double scaley;
        //double scalez;
        string SingleImageDir;
        string SequenceImageDir;
        string VideoDir;
        double duration;
        double cameraRatio;
        VideoTrack CameraVidtrack;
        Timecode cursorLocation;
        public SceneRatio sceneRatio;

        bool use_camera = true;

        public void Create_Initial()
        {
            sceneRatio = new SceneRatio();
            cameraRatio = 1;
            timeShift = 0;
        }

        public void Create_Setup()
        {
            Import_SetType();
            SetSceneRatio();
            SetCameraRatio();
            SetTimeShift();
        }

        public void Import_SetType()
        {

        }

        public void SetTimeShift()
        {
            // AE startTime is zero, CTA is 1
            // CTA's key time does not change by fps, so the devide value must be the default value 30
            cursorLocation = myVegas.Transport.CursorPosition;
            timeShift = Math.Round((double)((CTAScene.Create.startFrame - 1) * 1000) / 30) - cursorLocation.ToMilliseconds();  //timeshift in ms
        }

        public void Create_Project()
        {
            myCreate = CTAScene.Create;
            fps = CTAScene.fps;
            SceneName = CTAScene.name;
            SceneWidth = CTAScene.width;
            SceneHeight = CTAScene.height;
            scaletime = 1;
            //scalex = SceneWidth / 600;
            //offsetx = 300;
            //offsety = -75;
            //scaley = scalex;
            //scalez = 1;
            SingleImageDir = "Single Image";
            SequenceImageDir = "Sequence Image";
            VideoDir = "Video";
            duration = (myCreate.endFrame - myCreate.startFrame + 1) * 1000 / 30;
        }

        public void Comp_Create()
        {
            ShowLogMessage("---------------------------------------\r\nImport Started\r\n---------------------------------------\r\n");
            use_camera = true;
            // create parent track
            VideoTrack vidTrack = CreateFirstVideoTrack("Camera");
            cursorLocation = myVegas.Transport.CursorPosition;

            Comp_Add_Layer();
            if (CTAScene.smartType)
            {
                Comp_Add_Camera(vidTrack);
            }
            ShowLogMessage("---------------------------------------\r\nImport Finished\r\n---------------------------------------\r\n");
        }

        public void Audio_Create(string jsonComp_name, ref int trackindex)
        {
            string AudioDir = "Audio";
            string fileName;
            AudioTrack audioTrack;
            string s = "Create Audio: " + jsonComp_name;
            SaveLogFile(MethodBase.GetCurrentMethod(), s);

            fileName = Path.Combine(maindir, SceneName, AudioDir, jsonComp_name + ".wav");

            audioTrack = AddAudioTrack(jsonComp_name, trackindex);
            trackindex += 1;
            AddAudioFileToTimeline(fileName, audioTrack, cursorLocation);
        }

        public PosObj Layer_Interpolate_Position(PosObj prevpos, PosObj nextpos, int time_ms)
        {
            PosObj interpos = new PosObj();

            double time_factor = ((double)time_ms - (double)prevpos.time) / ((double)nextpos.time - (double)prevpos.time);

            interpos.x = prevpos.x + ((nextpos.x - prevpos.x) * time_factor);
            interpos.y = prevpos.y + ((nextpos.y - prevpos.y) * time_factor);
            interpos.z = prevpos.z + ((nextpos.z - prevpos.z) * time_factor);
            interpos.time = time_ms;

            return interpos;

        }

        public ScaleObj Layer_Interpolate_Scale(ScaleObj prevscale, ScaleObj nextscale, int time_ms)
        {
            ScaleObj interscale = new ScaleObj();

            double time_factor = ((double)time_ms - (double)prevscale.time) / ((double)nextscale.time - (double)prevscale.time);

            interscale.x = prevscale.x + ((nextscale.x - prevscale.x) * time_factor);
            interscale.y = prevscale.y + ((nextscale.y - prevscale.y) * time_factor);
            interscale.z = prevscale.z + ((nextscale.z - prevscale.z) * time_factor);
            interscale.time = time_ms;

            return interscale;

        }

        public PosObj Layer_Calculate_Position(PosObj posObj)
        {
            PosObj pos = new PosObj();

            double camRatio = this.cameraRatio;

            double transFactorX = sceneRatio.width / sceneRatio.scene;
            double transFactorY = sceneRatio.height / sceneRatio.scene;

            pos.x = transFactorX * posObj.x;
            pos.y = transFactorY * posObj.z;  // z and y are exchanged
            pos.z = posObj.y * camRatio;

            return pos;

        }

        public RotObj Layer_Calculate_Rotation(RotObj rotobj)
        {
            RotObj calcrot = new RotObj();

            double radian = 57.29577951;

            calcrot.x = 0; // rotobj.x * radian;
            calcrot.y = 0; //  rotobj.y * radian;
            calcrot.z = rotobj.y * radian;

            /*
             string s = "Layer_Calculate_Rotation\r\n";
            s += "calcrot.x  " + calcrot.x + "\r\n";
            s += "calcrot.y  " + calcrot.y + "\r\n";
            s += "calcrot.z  " + calcrot.z + "\r\n";

            //show_test_message(s);
            SaveLogFile(MethodBase.GetCurrentMethod(), s);
            */
            return calcrot;
        }

        public Pivot Layer_Calculate_AnchorPoint(Pivot pivot)
        {
            Pivot calcanchor = new Pivot();

            calcanchor.x = -pivot.x;

            calcanchor.y = pivot.y;

            calcanchor.z = pivot.z;

            return calcanchor;
        }

        public Pivot Layer_Search_Correponding_Pivot(Comp jsonComp, int time, ref int start_index)
        {
            Pivot pivot = new Pivot();
            Pivot search_pivot;
            Pivot last_search_pivot;
            double sceneRatio = this.sceneRatio.scene;
            last_search_pivot = jsonComp.compObj.pivot[0];

            for (int i = start_index; i < jsonComp.compObj.pivot.Count; i++)
            {
                search_pivot = jsonComp.compObj.pivot[i];
                if (search_pivot.time == time)
                {
                    pivot.x = search_pivot.x / sceneRatio;
                    pivot.y = search_pivot.z / sceneRatio;
                    pivot.z = search_pivot.y * this.cameraRatio;
                    start_index = i;
                    return pivot;
                }
                else
                {
                    if (search_pivot.time > time)
                    {
                        pivot.x = last_search_pivot.x / sceneRatio;
                        pivot.y = last_search_pivot.z / sceneRatio;
                        pivot.z = last_search_pivot.y * this.cameraRatio;
                        start_index = i - 1;
                        return pivot;
                    }
                }
                last_search_pivot = search_pivot;
            }
            pivot.x = last_search_pivot.x / sceneRatio;
            pivot.y = last_search_pivot.z / sceneRatio;
            pivot.z = last_search_pivot.y * this.cameraRatio;

            return pivot;
        }

        public ScaleObj Layer_Search_Correponding_Scale(Comp jsonComp, int time, int objidx)
        {
            ScaleObj scaleobj = new ScaleObj();
            double sceneRatio = this.sceneRatio.scene;

            ScaleObj search_scale = jsonComp.compObj.scaleObjs[objidx];
            {
                if (search_scale.time == time)
                {
                    scaleobj.x = search_scale.x;
                    scaleobj.y = search_scale.y;
                    scaleobj.z = search_scale.z;

                    return scaleobj;
                }
                else
                {
                    SaveLogFile(MethodBase.GetCurrentMethod(), "Error: Keyframes wrong time");
                }
            }

            scaleobj.x = 1;
            scaleobj.y = 1;
            scaleobj.z = 1;

            return scaleobj;
        }

        public RotObj Layer_Search_Correponding_RotObj(Comp jsonComp, int time, int objidx)
        {
            RotObj rotobj = new RotObj();

            RotObj search_rot = jsonComp.compObj.rotObjs[objidx];
            {
                if (search_rot.time == time)
                {
                    rotobj.x = search_rot.x;
                    rotobj.y = search_rot.y;
                    rotobj.z = search_rot.z;

                    return rotobj;
                }
                else
                {
                    SaveLogFile(MethodBase.GetCurrentMethod(), "Error: Keyframes wrong time");
                }
            }

            rotobj.x = 0;
            rotobj.y = 0;
            rotobj.z = 0;

            return rotobj;
        }



        public bool Layer_Add_PositionKey(ref VideoTrack vidTrack, Comp jsonComp)
        {
            CompObj compObj = jsonComp.compObj;
            SaveLogFile(MethodBase.GetCurrentMethod(), "started");

            ScaleObj scale1;
            ScaleObj calcScale = new ScaleObj();
            ScaleObj prevCalcScale = new ScaleObj();
            ScaleObj newCalcScale = new ScaleObj();
            ScaleObj compScale = new ScaleObj();
            compScale.x = jsonComp.compScale.scaleX;
            compScale.y = jsonComp.compScale.scaleZ;
            compScale.z = jsonComp.compScale.scaleY;
            double sceneRatio = this.sceneRatio.scene;
            int pivotstartindex = 0;
            int prevpivotindex = 0;
            int objidx = 0;
            Pivot calcPivot = new Pivot();
            //Pivot prevcalcPivot = new Pivot();
            Pivot pivot1;
            PosObj prevCalcPos = new PosObj();
            PosObj newCalcPos;
            PosObj calcpos;
            foreach (PosObj posobj in compObj.posObjs)
            {
                Timecode kftime = Timecode.FromMilliseconds(scaletime * posobj.time);
                TrackMotionKeyframe tmKF;
                if (posobj.time == -1)
                {
                    tmKF = vidTrack.TrackMotion.MotionKeyframes[0];
                }
                else
                {
                    tmKF = Search_Motion_Keyframe(ref vidTrack, scaletime * posobj.time);
                }

                calcpos = Layer_Calculate_Position(posobj);
                calcpos.time = posobj.time;

                prevpivotindex = pivotstartindex;
                pivot1 = Layer_Search_Correponding_Pivot(jsonComp, posobj.time, ref pivotstartindex);
                calcPivot = Layer_Calculate_AnchorPoint(pivot1);
                calcPivot.time = posobj.time;

                scale1 = Layer_Search_Correponding_Scale(jsonComp, posobj.time, objidx);
                calcScale = Layer_Calculate_Scale(scale1, compScale);
                calcScale.time = posobj.time;

                double factorWidth;
                double factorHeight;

                if ((orgWidth / orgHeight) > ((double)(SceneWidth) / (double)(SceneHeight))) // orgWidth is the scale factor
                {
                    factorWidth = 1;
                    factorHeight = orgHeight / orgWidth * SceneHeight / SceneWidth;

                    tmKF.Width = orgWidth * calcScale.x;
                    tmKF.Height = orgWidth * calcScale.y * SceneHeight / SceneWidth;
                }
                else                                       // objheight is the scale factor
                {
                    factorWidth = orgWidth / orgHeight * SceneWidth / SceneHeight;
                    factorHeight = 1;

                    tmKF.Width = orgHeight * calcScale.x * SceneWidth / SceneHeight;
                    tmKF.Height = orgHeight * calcScale.y;
                }

                double xfactor = (double)numParam07.Value / 100;
                double xoffset = (double)numParam01.Value;
                double yfactor = (double)numParam08.Value / 100;
                double yoffset = (double)numParam02.Value;
                double zfactor = (double)numParam09.Value / 100;
                double zoffset = (double)numParam03.Value;

                tmKF.Type = VideoKeyframeType.Linear;
                tmKF.PositionX = calcpos.x * xfactor - (calcPivot.x * calcScale.x) + xoffset;
                tmKF.PositionY = calcpos.y * yfactor + (calcPivot.y * calcScale.y) + yoffset;
                tmKF.PositionZ = calcpos.z * zfactor + (calcPivot.z * calcScale.z) + zoffset;

                RotObj rotobj = Layer_Search_Correponding_RotObj(jsonComp, posobj.time, objidx);
                RotObj calcRot = Layer_Calculate_Rotation(rotobj);

                tmKF.RotationX = calcRot.x;
                tmKF.RotationY = calcRot.y;
                tmKF.RotationZ = calcRot.z;

                tmKF.RotationOffsetX = +calcPivot.x * calcScale.x;
                tmKF.RotationOffsetY = -calcPivot.y * calcScale.y;
                tmKF.RotationOffsetZ = -calcPivot.z * calcScale.z;

                string s = "Layer_Pos_Pivot_Correction\r\n";
                s += jsonComp.name + "\r\n";
                s += posobj.time + "\r\n";
                s += "tmKF-Position" + tmKF.Position.ToPositionString() + "\r\n";
                s += "sceneRatio  " + sceneRatio + "\r\n";
                s += "calcPivot.x  " + calcPivot.x + "\r\n";
                s += "calcPivot.y  " + calcPivot.y + "\r\n";
                s += "calcPivot.z  " + calcPivot.z + "\r\n";
                s += "calcScale.x  " + calcScale.x + "\r\n";
                s += "calcScale.y  " + calcScale.y + "\r\n";
                s += "calcScale.z  " + calcScale.z + "\r\n";
                s += "calcpos.x  " + calcpos.x + "\r\n";
                s += "calcpos.y  " + calcpos.y + "\r\n";
                s += "calcpos.z  " + calcpos.z + "\r\n";
                s += "xfactor  " + xfactor + "\r\n";
                s += "xoffset  " + xoffset + "\r\n";
                s += "yfactor  " + yfactor + "\r\n";
                s += "yoffset  " + yoffset + "\r\n";
                s += "zfactor  " + zfactor + "\r\n";
                s += "zoffset  " + zoffset + "\r\n";
                s += "PositionX  " + tmKF.PositionX + "\r\n";
                s += "PositionY  " + tmKF.PositionY + "\r\n";
                s += "PositionZ  " + tmKF.PositionZ + "\r\n";
                s += "lastpivotindex  " + prevpivotindex + "\r\n";
                s += "pivotstartindex  " + pivotstartindex + "\r\n";

                string s1 = s;
                
                Pivot search_pivot;
                int time_ms;

                newCalcPos = calcpos;
                newCalcScale = calcScale;
               
                               
                // handle pivot keyframes between last and current poskeyframe
                if ((pivotstartindex-prevpivotindex)>1) // there are several pivot keyframes between lasxt pos keyframe and current keyframe
                {
                    for (int i = prevpivotindex + 1; i < pivotstartindex; i++) // handle all pivot keyframes between last pivot index and current pivot index 
                    {
                        search_pivot = jsonComp.compObj.pivot[i];
                        pivot1.x = search_pivot.x / sceneRatio;
                        pivot1.y = search_pivot.z / sceneRatio;
                        pivot1.z = search_pivot.y; // * this.cameraRatio;

                        calcPivot = Layer_Calculate_AnchorPoint(pivot1);
                        time_ms = search_pivot.time;
                        calcScale = Layer_Interpolate_Scale(prevCalcScale, newCalcScale, time_ms); //prevCalcScale; // possible interpolation between current calcScale and lastCalcScale needs to be added

                        //kftime = Timecode.FromMilliseconds(time_ms);
                        //tmKF = vidTrack.TrackMotion.InsertMotionKeyframe(kftime);
                        //s = "new keyframe at: " + tmKF.Position.ToPositionString();
                        //SaveLogFile(MethodBase.GetCurrentMethod(), s);

                        tmKF = Search_Motion_Keyframe(ref vidTrack, time_ms);

                        tmKF.Type = VideoKeyframeType.Linear;
                        // for position the interpolation calculated by VEGAS is used, therefore only the pivot changes are added

                        Pivot deltapivot = new Pivot();

                        calcpos = Layer_Interpolate_Position(prevCalcPos,newCalcPos,time_ms);

                        tmKF.Type = VideoKeyframeType.Linear;

                        tmKF.PositionX = calcpos.x * xfactor - (calcPivot.x * calcScale.x) + xoffset;
                        tmKF.PositionY = calcpos.y * yfactor + (calcPivot.y * calcScale.y) + yoffset;
                        tmKF.PositionZ = calcpos.z * zfactor + (calcPivot.z * calcScale.z) + zoffset;

                        tmKF.RotationOffsetX = +calcPivot.x * calcScale.x;
                        tmKF.RotationOffsetY = -calcPivot.y * calcScale.y;
                        tmKF.RotationOffsetZ = -calcPivot.z * calcScale.z;

                        s = "Layer_Pos_Pivot_Correction Additional keyframe\r\n";
                        s += jsonComp.name + "\r\n";
                        s += "Pivot-Time ms:"+ time_ms + "\r\n";
                        s += "tmKF-Position" + tmKF.Position.ToPositionString() + "\r\n";
                        s += "search_pivot.time" + search_pivot.time + "\r\n";
                        s += "search_pivot.x  " + search_pivot.x + "\r\n";
                        s += "search_pivot.y  " + search_pivot.y + "\r\n";
                        s += "search_pivot.z  " + search_pivot.z + "\r\n";
                        s += "calcpivot.x  " + calcPivot.x + "\r\n";
                        s += "calcpivot.y  " + calcPivot.y + "\r\n";
                        s += "calcpivot.z  " + calcPivot.z + "\r\n";
                        s += "scale.x  " + calcScale.x + "\r\n";
                        s += "scale.y  " + calcScale.y + "\r\n";
                        s += "scale.z  " + calcScale.z + "\r\n";
                        s += "calcpos.x  " + calcpos.x + "\r\n";
                        s += "calcpos.y  " + calcpos.y + "\r\n";
                        s += "calcpos.z  " + calcpos.z + "\r\n";
                        s += "xfactor  " + xfactor + "\r\n";
                        s += "xoffset  " + xoffset + "\r\n";
                        s += "yfactor  " + yfactor + "\r\n";
                        s += "yoffset  " + yoffset + "\r\n";
                        s += "zfactor  " + zfactor + "\r\n";
                        s += "zoffset  " + zoffset + "\r\n";
                        s += "PositionX  " + tmKF.PositionX + "\r\n";
                        s += "PositionY  " + tmKF.PositionY + "\r\n";
                        s += "PositionZ  " + tmKF.PositionZ + "\r\n";
                        s += "RotationOffsetX  " + tmKF.RotationOffsetX + "\r\n";
                        s += "RotationOffsetY  " + tmKF.RotationOffsetY + "\r\n";
                        s += "RotationOffsetZ  " + tmKF.RotationOffsetZ + "\r\n";
                        SaveLogFile(MethodBase.GetCurrentMethod(), s);
                    }
  
                }
                SaveLogFile(MethodBase.GetCurrentMethod(), s1); // log current keyframe data
                prevCalcPos = newCalcPos;
                prevCalcScale = newCalcScale;
                objidx++;
            }

            return true;
        }

        public bool Layer_Add_AnchorPointKey(ref VideoTrack vidTrack, Comp jsonComp)
        {
            return true;
            /*
            CompObj compObj = jsonComp.compObj;

            PosObj imageCenter = new PosObj();

            imageCenter.x = orgWidth / 2;
            imageCenter.y = orgHeight / 2;
            imageCenter.z = 0;

            foreach (Pivot pivot in compObj.pivot)
            {
                Timecode kftime = Timecode.FromMilliseconds(scaletime * pivot.time);
                TrackMotionKeyframe tmKF;
                if (pivot.time == -1)
                {
                    tmKF = vidTrack.TrackMotion.MotionKeyframes[0];
                }
                else 
                {
                    tmKF = Search_Motion_Keyframe(ref vidTrack, scaletime * pivot.time);
                    //tmKF = vidTrack.TrackMotion.InsertMotionKeyframe(kftime);
                }

                Pivot calcanchor = Layer_Calculate_AnchorPoint(pivot, imageCenter);

                tmKF.RotationOffsetX = calcanchor.x;
                tmKF.RotationOffsetY = calcanchor.y;
                tmKF.RotationOffsetY = calcanchor.z;

            }
            return true;
            */
        }

        public int lastkey_index = 0;

        public TrackMotionKeyframe Search_Motion_Keyframe(ref VideoTrack vidTrack, double time)
        {
            string s;
            time = scaletime * time - timeShift;
            if (time == cursorLocation.ToMilliseconds())
            {
                lastkey_index = 0;
            }
            for (int i = lastkey_index; i < vidTrack.TrackMotion.MotionKeyframes.Count; i++)
            {
                TrackMotionKeyframe tmkf = vidTrack.TrackMotion.MotionKeyframes[i];
                lastkey_index = i;

                if (tmkf.Position.ToMilliseconds() == time)
                {
                    s = "Keyframe found: " + tmkf.Position.ToPositionString();
                    SaveLogFile(MethodBase.GetCurrentMethod(), s);
                    return tmkf;
                }
                else
                {
                    if (tmkf.Position.ToMilliseconds() > time)
                    {
                        break;
                    }
                }
            }
            Timecode kftime = Timecode.FromMilliseconds(time);
            TrackMotionKeyframe tmkf1 = vidTrack.TrackMotion.InsertMotionKeyframe(kftime);
            s = "Keyframe not found - new keyframe at: " + tmkf1.Position.ToPositionString();
            SaveLogFile(MethodBase.GetCurrentMethod(), s);
            return tmkf1;
        }

        public TrackMotionKeyframe Search_ParentMotion_Keyframe(ref VideoTrack vidTrack, double time)
        {
            string s;
            time = time - timeShift;
            foreach (TrackMotionKeyframe tmkf in vidTrack.ParentTrackMotion.MotionKeyframes)
            {
                if (tmkf.Position.ToMilliseconds() == time)
                {
                    return tmkf;
                }
            }
            Timecode kftime = Timecode.FromMilliseconds(scaletime * time);
            TrackMotionKeyframe tmkf1 = vidTrack.ParentTrackMotion.InsertMotionKeyframe(kftime);
            return tmkf1;
        }

        public ScaleObj Layer_Calculate_Scale(ScaleObj compscaleObj, ScaleObj compScale)
        {
            ScaleObj calcscale = new ScaleObj();

            double scaleFactorX = compScale.x * sceneRatio.width; // * 100;
            double scaleFactorY = compScale.y * sceneRatio.height; // * 100;

            calcscale.x = compscaleObj.x * scaleFactorX;
            calcscale.y = compscaleObj.z * scaleFactorY;
            calcscale.z = 1;

            string s = "compScale.x   " + compScale.x + "\r\n";
            s = s + "sceneRatio.width  " + sceneRatio.width + "\r\n";
            s += "compScale.y  " + compScale.y + "\r\n";
            s += "sceneRatio.height  " + sceneRatio.height + "\r\n";
            s += "scaleFactorX  " + scaleFactorX + "\r\n";
            s += "scaleFactorY  " + scaleFactorY + "\r\n";
            s += "compscaleObj.x  " + compscaleObj.x + "\r\n";
            s += "compscaleObj.z  " + compscaleObj.z + "\r\n";
            s += "calcscale.x  " + calcscale.x + "\r\n";
            s += "calcscale.y  " + calcscale.y + "\r\n";
            SaveLogFile(MethodBase.GetCurrentMethod(), s);

            return calcscale;
        }

        public bool Layer_Add_ScaleKey(ref VideoTrack vidTrack, Comp jsonComp)
        {
            /*
            
            CompObj compObj = jsonComp.compObj;

            ScaleObj compScale = new ScaleObj();
            compScale.x = jsonComp.compScale.scaleX;
            compScale.y = jsonComp.compScale.scaleZ;
            compScale.z = jsonComp.compScale.scaleY;

            string s;

            foreach (ScaleObj scaleobj in compObj.scaleObjs)
            {
                Timecode kftime = Timecode.FromMilliseconds(scaletime * scaleobj.time);
                TrackMotionKeyframe tmKF;
                if (scaleobj.time == -1)
                {
                    tmKF = vidTrack.TrackMotion.MotionKeyframes[0];
                }
                else
                {
                    tmKF = Search_Motion_Keyframe(ref vidTrack, scaletime * scaleobj.time);
                }
                tmKF.Type = VideoKeyframeType.Linear;

                ScaleObj calcScale = Layer_Calculate_Scale(scaleobj, compScale);

                if ((orgWidth / orgHeight) > ((double)(SceneWidth) / (double)(SceneHeight))) // orgWidth is the scale factor
                {
                    s = "\r\n";
                    s += "Objwidth" + "\r\n";
                    s += "SceneWidth" + SceneWidth + "\r\n";
                    s += "SceneHeight" + SceneHeight + "\r\n";
                    s += "SceneWidth/SceneHeight" + SceneWidth / SceneHeight + "\r\n";
                    SaveLogFile(MethodBase.GetCurrentMethod(), s);

                    tmKF.Width = orgWidth * calcScale.x;
                    tmKF.Height = orgWidth * calcScale.y * SceneHeight / SceneWidth;
                }
                else                                       // objheight is the scale factor
                {
                    s = "\r\n";
                    s += "Objheight" + "\r\n";
                    s += "SceneWidth" + SceneWidth + "\r\n";
                    s += "SceneHeight" + SceneHeight + "\r\n";
                    s += "SceneWidth/SceneHeight" + SceneWidth / SceneHeight + "\r\n";
                    s += "calcScale.x   " + calcScale.x + "\r\n";
                    s += "calcScale.y   " + calcScale.y + "\r\n";
                    SaveLogFile(MethodBase.GetCurrentMethod(), s);

                    tmKF.Width = orgHeight * calcScale.x * SceneWidth / SceneHeight;
                    tmKF.Height = orgHeight * calcScale.y;
                }

                s = "\r\n";
                s += jsonComp.name + "\r\n";
                s += scaleobj.time + "\r\n";
                s += "tmKF-Position" + tmKF.Position.ToPositionString() + "\r\n";
                s += "orgWidth   " + orgWidth + "\r\n";
                s += "calcScale.x  " + calcScale.x + "\r\n";
                s += "tmKF.Width  " + tmKF.Width + "\r\n";
                s += "sceneRatio.height  " + sceneRatio.height + "\r\n";
                s += "orgHeight  " + orgHeight + "\r\n";
                s += "calcScale.y  " + calcScale.y + "\r\n";
                s += "tmKF.Height  " + tmKF.Height + "\r\n";

                SaveLogFile(MethodBase.GetCurrentMethod(), s);
            }
            */
            return true;
        }

        public bool Layer_Add_RotationKey(ref VideoTrack vidTrack, Comp jsonComp)
        {
            /*
            CompObj compObj = jsonComp.compObj;

            foreach (RotObj rotobj in compObj.rotObjs)
            {
                Timecode kftime = Timecode.FromMilliseconds(scaletime * rotobj.time);
                TrackMotionKeyframe tmKF;
                if (rotobj.time == -1)
                {
                    tmKF = vidTrack.TrackMotion.MotionKeyframes[0];
                }
                else
                {
                    tmKF = Search_Motion_Keyframe(ref vidTrack, scaletime * rotobj.time);
                }
                tmKF.Type = VideoKeyframeType.Linear;

                RotObj calcRot = Layer_Calculate_Rotation(rotobj);

                tmKF.RotationX = calcRot.x;
                tmKF.RotationY = calcRot.y;
                tmKF.RotationZ = calcRot.z;

            }
            */
            return true;
        }


        public void Layer_SetProperty(ref VideoTrack vidTrack, Comp jsonComp)
        {
            if (CTAScene.smartType)
            {
                if ((Layer_Add_PositionKey(ref vidTrack, jsonComp) &&
                    Layer_Add_AnchorPointKey(ref vidTrack, jsonComp) &&
                    Layer_Add_RotationKey(ref vidTrack, jsonComp) &&
                    Layer_Add_ScaleKey(ref vidTrack, jsonComp))
                    == false)
                {
                    //error message
                }
            }
        }

        public void Layer_Create(Comp jsonComp, ref int trackindex,bool updateonly=false)
        {
            string ImageDir;
            string fileName;
            VideoTrack vidTrack;

            CompObj compObj = jsonComp.compObj;

            if (updateonly == false)
            {

                if (jsonComp.hasVideo)
                {
                    ImageDir = VideoDir;
                    fileName = Path.Combine(maindir, SceneName, ImageDir, compObj.fileName);
                }
                else
                {
                    if (compObj.count == 1)
                    {
                        ImageDir = SingleImageDir;
                        fileName = Path.Combine(maindir, SceneName, ImageDir, compObj.fileName);
                    }
                    else
                    {
                        ImageDir = SequenceImageDir;
                        fileName = Path.Combine(maindir, SceneName, ImageDir, jsonComp.name, compObj.fileName);
                    }
                }

                vidTrack = AddVideoTrack(jsonComp.name, trackindex);

                trackindex += 1;
                AddFileToTimeline(fileName, vidTrack, cursorLocation, compObj.count, fps);
                vidTrack.CompositeMode = CompositeMode.SrcAlpha3D;
                vidTrack.CompositeNestingLevel = 1;
            }
            else
            {
                vidTrack = (VideoTrack) myVegas.Project.Tracks[trackindex];
             
                string s = "\r\n";
                s += "Update Track " + trackindex + "\r\n";
                s += "Trackname " + vidTrack.Name + "\r\n";
                SaveLogFile(MethodBase.GetCurrentMethod(), s);

                trackindex += 1;
            }

            Layer_SetProperty(ref vidTrack, jsonComp);

        }


        // ***** Camera *******
        public void Camera_Create(VideoTrack vidTrack, Camera jsonCamera, int trackindex)
        {
            SaveLogFile(MethodBase.GetCurrentMethod(), "Started");
            AECamera aeCamera = new AECamera();
            aeCamera.name = "AECamera";
            aeCamera.pos = new PosObj();
            aeCamera.pos.x = CTAScene.width / 2;
            aeCamera.pos.y = CTAScene.height / 2;

            // set value zoom by focal length
            double value = jsonCamera.camLens * cameraFLM;
            aeCamera.zoom = value;
            aeCamera.focusDistance = value;

            Camera_SetProperty(vidTrack, aeCamera, jsonCamera);
        }

        public void Camera_SetProperty(VideoTrack vidTrack, AECamera aeCamera, Camera jsonCamera)
        {
            SaveLogFile(MethodBase.GetCurrentMethod(), "Started");
            if (false == (Camera_Add_PositionKey(vidTrack, aeCamera, jsonCamera)
                         && Camera_Add_RotateKey(vidTrack, aeCamera, jsonCamera)))
            {

                //error message
            }
        }

        public bool Camera_Add_PositionKey(VideoTrack vidTrack, AECamera aeCamera, Camera jsonCamera)
        {
            SaveLogFile(MethodBase.GetCurrentMethod(), "Started");

            double value = jsonCamera.camLens * cameraFLM;

            foreach (PosObj posobj in jsonCamera.posObjs)
            {
                Timecode kftime = Timecode.FromMilliseconds(scaletime * posobj.time);
                TrackMotionKeyframe tmKF;
                vidTrack.ParentCompositeMode = CompositeMode.SrcAlpha3D;
                if (posobj.time == 0)
                {
                    tmKF = vidTrack.ParentTrackMotion.MotionKeyframes[0];
                }
                else
                {
                    tmKF = Search_ParentMotion_Keyframe(ref vidTrack, scaletime * posobj.time);
                }
                PosObj calcpos = Camera_Calculate_Position(posobj);

                PosObj posfactor = new PosObj();
                posfactor.x = -(double)numParam10.Value / 100;
                posfactor.y = (double)numParam11.Value / 100;
                posfactor.z = (double)numParam12.Value / 100;

                PosObj posoffset = new PosObj();
                posoffset.x = (double)numParam04.Value;
                posoffset.y = (double)numParam05.Value;
                posoffset.z = (double)numParam06.Value;

                tmKF.Type = VideoKeyframeType.Linear;
                tmKF.PositionX = (calcpos.x * posfactor.x) + posoffset.x;
                tmKF.PositionY = (calcpos.y * posfactor.y) + posoffset.y;
                tmKF.PositionZ = (calcpos.z - 1080) * posfactor.z + posoffset.z;

                orgWidth = (int)((double)SceneWidth);
                if (orgWidth > 0)
                {
                    tmKF.Width = orgWidth;
                    tmKF.Height = orgWidth;
                }

                tmKF.Depth = SceneHeight;

                string s = "Camera_add_Pos_key\r\n";
                s += "Camera\r\n";
                s += posobj.time + "\r\n";
                s += "calcpos.x  " + calcpos.x + "\r\n";
                s += "calcpos.y  " + calcpos.y + "\r\n";
                s += "calcpos.z  " + calcpos.z + "\r\n";
                s += "xfactor  " + posfactor.x + "\r\n";
                s += "xoffset  " + posoffset.x + "\r\n";
                s += "yfactor  " + posfactor.y + "\r\n";
                s += "yoffset  " + posoffset.y + "\r\n";
                s += "zfactor  " + posfactor.z + "\r\n";
                s += "zoffset  " + posoffset.y + "\r\n";

                SaveLogFile(MethodBase.GetCurrentMethod(), s);
            }
            return true;
        }

        public bool Camera_Add_RotateKey(VideoTrack vidTrack, AECamera aeCamera, Camera jsonCamera)
        {
            /*
            var cameraRot = aeCamera.zRotation;

            var rotKeys = Camera_Calculate_Rotation(jsonCamera);

            for (var i = 0; i < rotKeys.length; ++i)
            {

                var zRot = rotKeys[i].value;
                var keyTime = rotKeys[i].time;

                try
                {

                    var index = cameraRot.addKey(keyTime);
                    cameraRot.setValueAtKey(index, zRot);
                }
                catch (e)
                {

            alert( e.message );
                    return false;
                }
            }
            */
            return true;
        }


        public PosObj Camera_Calculate_Position(PosObj posObj)
        {
            PosObj pos = new PosObj();

            double camRatio = -this.cameraRatio;

            double transFactorX = sceneRatio.width / sceneRatio.scene;
            double transFactorY = sceneRatio.height / sceneRatio.scene;

            //double keyTime = Math.Round((pos.time - timeShift) * fps * 1000) / 1000;
            //pos.time = Math.Round(keyTime) / fps;

            pos.x = transFactorX * (posObj.x);
            pos.y = transFactorY * (-posObj.z);  // z and y are exchanged
            pos.z = posObj.y * camRatio; // (-0.995);

            string s = "Camera_Calculate_Position\r\n";
            s += "Camera\r\n";
            s += posObj.time + "\r\n";
            s += "transFactorX  " + transFactorX + "\r\n";
            s += "transFactorY  " + transFactorY + "\r\n";
            s += "camRatio  " + camRatio + "\r\n";
            s += "pos.x  " + pos.x + "\r\n";
            s += "pos.y  " + pos.y + "\r\n";
            s += "pos.z  " + pos.z + "\r\n";

            SaveLogFile(MethodBase.GetCurrentMethod(), s);

            return pos;

        }

        public int Camera_Calculate_Rotation(RotObj rotObj)
        {
            int value = -rotObj.value;
            return value;
        }

        public void ImportCTAProject(Vegas vegas)
        {
            myVegas = vegas;
            SaveLogFile(MethodBase.GetCurrentMethod(), "Started", true);

            bool stopexecution = false;

            FileName = GetFileName("Please select CartoonAnimator JSON File",FileName);

            if (FileName != null)
            {
                ShowLogMessage("Opening JSON file " + FileName + "\r\n");
                Create_Initial();
                try
                {
                    jsonString = File.ReadAllText(FileName);
                }
                catch (Exception e)
                {
                    ShowLogMessage("**File-ERROR** " + e.Message + "\r\n");
                    stopexecution=true;
                }

                try
                {
                    CTAScene = JsonSerializer.Deserialize<Root>(jsonString);
                }
                catch (Exception e)
                {
                    ShowLogMessage("**JSON-ERROR** " + e.Message + "\r\n");
                    stopexecution=true;
                }
                if (!stopexecution)
                {
                    maindir = Path.GetDirectoryName(FileName);
                    Create_Project();
                    Create_Setup();
                    Comp_Create();
                }
            }
        }

        public void Comp_Add_Layer()
        {
            int trackindex = 1;

            SaveLogFile(MethodBase.GetCurrentMethod(), "Started");

            foreach (Comp jsonComp in myCreate.comps)
            {
                string s = "------------------------------------------------------------------------------------------------" + "\r\n";
                s += "* Add Layer started: " + jsonComp.name + "\r\n";
                s += "* HasAudio:" + jsonComp.hasAudio + "\r\n";
                s += "* HasVideo:" + jsonComp.hasVideo + "\r\n";
                s += "------------------------------------------------------------------------------------------------" + "\r\n";
                SaveLogFile(MethodBase.GetCurrentMethod(), s);
                ShowLogMessage("Create Track: " + jsonComp.name + "\r\n");

                if (jsonComp.hasAudio == true)
                {
                    Audio_Create(jsonComp.name, ref trackindex);
                }
                Layer_Create(jsonComp, ref trackindex);
            }
        }

        public void Comp_Add_Camera(VideoTrack vidTrack)
        {
            Camera jsonCamera = myCreate.camera;
            /*
            if (isEmptyObject(jsonCamera))
            {
            
                return;
            }
            */
            ShowLogMessage("Create Cameratrack\r\n");
            if (use_camera == true)
            {
                CameraVidtrack = vidTrack;

                Camera_Create(vidTrack, jsonCamera, 0);
            }

        }

        public void SetSceneRatio()
        {
            sceneRatio.scene = CTAScene.Create.sceneRatio;

            double sceneWidth = Math.Abs(CTAScene.viewCornerLT.x - CTAScene.viewCornerRB.x) / sceneRatio.scene;
            double sceneHeight = Math.Abs(CTAScene.viewCornerLT.z - CTAScene.viewCornerRB.z) / sceneRatio.scene;
            double widthRatio = CTAScene.width / sceneWidth;
            double heightRatio = CTAScene.height / sceneHeight;

            string s = "sceneWidth   " + sceneWidth + "\r\n";
            s = s + "sceneHeight  " + sceneHeight + "\r\n";
            s += "CTAScene.width  " + CTAScene.width + "\r\n";
            s += "CTAScene.height  " + CTAScene.height + "\r\n";
            s += "widthRatio  " + widthRatio + "\r\n";
            s += "heightRatio  " + heightRatio + "\r\n";
            s += "sceneRatio.scene  " + sceneRatio.scene + "\r\n";


            SaveLogFile(MethodBase.GetCurrentMethod(), s);

            sceneRatio.width = widthRatio;
            sceneRatio.height = heightRatio;
        }

        public void SetCameraRatio()
        {
            var value = CTAScene.Create.camera.camLens * cameraFLM;
            this.cameraRatio = (Math.Abs(value / standardZinCTACam)) * 0.26 * (double)numParam13.Value/100.0;

            string s = "cameraRatio   " + this.cameraRatio + "\r\n";
 
            SaveLogFile(MethodBase.GetCurrentMethod(), s);

            return ;
        }
 
        public string GetFileName(string Title, string filename)
        {
            OpenFileDialog svd = new OpenFileDialog();
            svd.Title = Title;
            if (this.debug == true)
            {
                svd.InitialDirectory = "D:\\CartoonAnimator\\AE_export";
            }
            if (filename != "")
            {
                svd.FileName = filename;
            }

            if (svd.ShowDialog() == DialogResult.OK)
            { return svd.FileName; }
            else
            {
                MessageBox.Show("No file was selected");
                return null;
            }
        }

        public VideoTrack CreateFirstVideoTrack(string trackname)
        {
            VideoTrack vidTrack = null;
            VideoTrack newTrack = new VideoTrack(myVegas.Project, 0, trackname);
            myVegas.Project.Tracks.Add(newTrack);
            vidTrack = newTrack;
            vidTrack.CompositeMode = CompositeMode.SrcAlpha3D;
            return vidTrack;
        }

        public VideoTrack AddVideoTrack(string trackname, int trackindex)
        {
            VideoTrack vidTrack = null;
            VideoTrack newTrack = new VideoTrack(myVegas.Project, trackindex, trackname);
            myVegas.Project.Tracks.Add(newTrack);
            vidTrack = newTrack;
            return vidTrack;
        }

        public AudioTrack AddAudioTrack(string trackname, int trackindex)
        {
            AudioTrack newTrack = new AudioTrack(myVegas.Project, trackindex, trackname);
            myVegas.Project.Tracks.Add(newTrack);
            return newTrack;
        }

        public void AddFileToTimeline(string FileName, VideoTrack vidTrack, Timecode position, int count, double fps)
        {
            Media mymedia = null;
            orgWidth = 0;
            orgHeight = 0;

            if (count == 1)
            {
                mymedia = Media.CreateInstance(myVegas.Project, FileName);
                string s = "Filename: " + FileName + "\r\n";
                SaveLogFile(MethodBase.GetCurrentMethod(), s);
                ShowLogMessage("Added " + s);
            }
            else
            {
                mymedia = myVegas.Project.MediaPool.AddImageSequence(FileName, count, fps);
                string s = FileName + "(" + count + ")" + "\r\n";
                SaveLogFile(MethodBase.GetCurrentMethod(), s);
                ShowLogMessage("Added ImageSequence:" + s);
            }

            if (mymedia.HasVideo())
            {
                string s = "Has Video" + "\r\n";
                SaveLogFile(MethodBase.GetCurrentMethod(), s);

                MediaStream stream = mymedia.Streams.GetItemByMediaType(MediaType.Video, 0);
                if (stream != null)
                {
                    Timecode length_tc;
                    if (count == 0) // Image only
                    {
                        int length_ms = (int)(duration);
                        length_tc = Timecode.FromMilliseconds(length_ms);
                        s = "Count=" + count + "\r\n";
                        s += "Duration: (s) " + duration + "\r\n";
                        SaveLogFile(MethodBase.GetCurrentMethod(), s);
                    }
                    else
                    {
                        s = "Count=" + count + "\r\n";
                        s += "Duration: (pos)" + stream.Length.ToPositionString() + "\r\n";
                        SaveLogFile(MethodBase.GetCurrentMethod(), s);
                        if (stream.Length.ToMilliseconds() > 0)
                        {
                            length_tc = stream.Length;
                        }
                        else
                        {
                            int length_ms = (int)(duration);
                            length_tc = Timecode.FromMilliseconds(length_ms);
                        }

                    }
                    s = "Create VideoEvent - Count=" + count + "\r\n";
                    s += "Eventlength: " + length_tc.ToPositionString() + "\r\n";
                    SaveLogFile(MethodBase.GetCurrentMethod(), s);
                    VideoEvent myNewVideoEvent = new VideoEvent(position, length_tc);
                    vidTrack.Events.Add(myNewVideoEvent);
                    Take myNewTake = new Take(stream);
                    myNewVideoEvent.Takes.Add(myNewTake);
                    VideoStream vs = (VideoStream)stream;
                    orgWidth = vs.Width;
                    orgHeight = vs.Height;
                }
            }
        }

        public void AddAudioFileToTimeline(string FileName, AudioTrack audioTrack, Timecode position)
        {
            Media mymedia = null;
            orgWidth = 0;
            orgHeight = 0;

            string s = "Has Audio" + "\r\n";
            SaveLogFile(MethodBase.GetCurrentMethod(), s);
            mymedia = Media.CreateInstance(myVegas.Project, FileName);
            s = "Create Media - Filename: " + FileName + "\r\n";
            ShowLogMessage(s);
            SaveLogFile(MethodBase.GetCurrentMethod(), s);

            MediaStream stream = mymedia.Streams.GetItemByMediaType(MediaType.Audio, 0);
            if (stream != null)
            {
                s = "Create AudioEvent " + FileName + "\r\n";
                SaveLogFile(MethodBase.GetCurrentMethod(), s);
                AudioEvent myNewAudioEvent = new AudioEvent(position, stream.Length);
                audioTrack.Events.Add(myNewAudioEvent);
                Take myNewTake = new Take(stream);
                myNewAudioEvent.Takes.Add(myNewTake);
            }
        }



        public void SaveLogFile(object method, string message, bool start = false)
        {
            if (debug == true)
            {
                string location = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) + @"\VEGAS Pro\20.0\";
                FileMode mode;
                if (start == true)
                {
                    mode = FileMode.Create;
                }
                else
                {
                    mode = FileMode.Append;
                }

                try
                {
                    //Opens a new file stream which allows asynchronous reading and writing
                    using (StreamWriter sw = new StreamWriter(new FileStream(location + @"CTAImport_log.txt", mode, FileAccess.Write, FileShare.ReadWrite)))
                    {
                        //Writes the method name with the exception and writes the exception underneath
                        sw.WriteLine(String.Format("{0} ({1}) - Method: {2} {3}", DateTime.Now.ToShortDateString(), DateTime.Now.ToShortTimeString(), method.ToString(), message));
                    }
                }
                catch (IOException)
                {
                    if (!File.Exists(location + @"log.txt"))
                    {
                        File.Create(location + @"log.txt");
                    }
                }
            }

        }

        public void ShowLogMessage(string message, bool start = false)
        {
            if (start == true)
            {
                LogTextbox.Text = "";
            }
            LogTextbox.AppendText(message);

        }

        private void numParam01_ValueChanged(object sender, EventArgs e)
        {

        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {

        }

        private void numParam07_ValueChanged(object sender, EventArgs e)
        {

        }
    }
}
 

