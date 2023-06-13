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
using System.Windows.Forms;
using ScriptPortal.Vegas;
using ScriptPortal.MediaSoftware.Skins;

namespace CTAImportTest
{
    public partial class CCMainForm : UserControl
    {
        Vegas myVegas;

        //public Vegas myVegas;

        public CCMainForm(Vegas vegas)
        {
            myVegas = vegas;
            SaveLogFile(MethodBase.GetCurrentMethod(), "Started");

            InitializeComponent();
            this.BackColor = lblMain1.BackColor;
            groupBox1.ForeColor = lblMain1.ForeColor;
            lblMain1.Text = "Import CartoonAnimator projects into VEGAS";
            lblMain2.Text = "In order to link an animated project into VEGAS, kindly import the ";
            lblMain3.Text = "Cartoon Animator JSON file created by the <Export to After Effects> command";

            checkBox1.CheckState = CheckState.Checked;
 
            lblParam01.Text = "Offset x";
            lblParam02.Text = "Offset y";
            lblParam03.Text = "Offset z";
            lblParam04.Text = "Pivot x factor";
            lblParam04.Visible = false;
            lblParam05.Text = "Pivot Y Factor";
            lblParam05.Visible = false;
            lblParam06.Text = "";
            lblParam06.Visible = false;
            lblParam07.Text = "Stretch Factor x";
            lblParam08.Text = "Stretch Factor y";
            lblParam09.Text = "Stretch Factor z";
            lblParam10.Text = "camdepth";
            lblParam10.Visible = false;
            lblParam11.Text = "camwindow-Width";
            lblParam11.Visible = false;
            lblParam12.Text = "";
            lblParam12.Visible = false;
            lblParam13.Text = "CameraRatio-Factor";
            lblParam13.Visible = false;
            lbl_copyright.Text = "© Harold Linke 2023 - Version 0.04";
            groupBox1.Text = "Finetune Camera Keyframes";


            numParam01.Value = 0;
            numParam02.Value = 0;
            numParam03.Value = 0;
            numParam04.Value = 100;
            numParam04.Visible = false;
            numParam05.Value = 100;
            numParam05.Visible = false;
            numParam06.Value = 100;
            numParam06.Visible = false;
            numParam07.Value = 100;
            numParam08.Value = 100;
            numParam09.Value = 100;
            numParam10.Value = 100;
            numParam10.Visible = false;
            numParam11.Value = 100;
            numParam11.Visible = false;
            numParam12.Value = 0;
            numParam12.Visible = false;
            numParam13.Value = 100;
            numParam13.Visible = false;

            checkBox1.Visible = false;
            checkBox3.Visible = false;

        }
        
        public void btnImportCTA_Click(object sender, EventArgs e)
        {
            string CTAVersion = "0.01";
            string s = "CTA Import - Version:" + CTAVersion;
            SaveLogFile(MethodBase.GetCurrentMethod(), "Started",true);
            using (UndoBlock undo = new UndoBlock("Import CTA Project"))
            {
                ImportCTAProject(myVegas);
            }
        }

        public void btnRecalcCameraKeyFrames_Click(object sender, EventArgs e)
        {
            using (UndoBlock undo = new UndoBlock("Recalc Camera KeyFrames"))
            {
                SaveLogFile(MethodBase.GetCurrentMethod(), "Started");
                Camera_Create(CameraVidtrack, myCreate.camera, 0); ;
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

        string maindir;

        public Create myCreate;
        double fps;
        string SceneName;
        int SceneWidth;
        int SceneHeight;
        double scaletime;
        double scalex;
        double offsetx;
        double offsety;
        double scaley;
        double scalez;
        string SingleImageDir;
        string SequenceImageDir;
        string VideoDir;
        double duration;
        double cameraRatio;
        VideoTrack CameraVidtrack;
        Timecode cursorLocation;
        public SceneRatio sceneRatio;

        bool use_pivot = false;
        bool use_camera = true;

        public void show_test_message(string s)
        {
            bool test = checkBox3.Checked;

            if (test == true)
                MessageBox.Show(s);
        }

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
            timeShift = Math.Round((double)((CTAScene.Create.startFrame - 1) / 30) * 1000) / 1000;
        }

        public void Create_Project()
        {
            myCreate = CTAScene.Create;
            fps = CTAScene.fps;
            SceneName = CTAScene.name;
            SceneWidth = CTAScene.width;
            SceneHeight = CTAScene.height;
            scaletime = 1;
            scalex = SceneWidth / 600;
            offsetx = 300;
            offsety = -75;
            scaley = scalex; // SceneHeight/600;
            scalez = 1;
            SingleImageDir = "Single Image";
            SequenceImageDir = "Sequence Image";
            VideoDir = "Video";
            duration = (myCreate.endFrame - myCreate.startFrame + 1) * 1000 / 30; // error in CTA, endframe is based on 30 FPS and not on real fps - CTAScene.fps;
        }

        public void Comp_Create()
        {
            use_camera = checkBox1.Checked;
            // create parent track
            VideoTrack vidTrack = CreateFirstVideoTrack("Camera");
            cursorLocation = myVegas.Transport.CursorPosition;

            Comp_Add_Layer();


            if (CTAScene.smartType)
            {
                Comp_Add_Camera(vidTrack);
            }
        }

        public void Audio_Create(string jsonComp_name,ref int trackindex)
        {
            string AudioDir = "Audio";
            string fileName;
            AudioTrack audioTrack;
            string s = "Create Audio: " + jsonComp_name;
            SaveLogFile(MethodBase.GetCurrentMethod(), s);

            fileName = Path.Combine(maindir, SceneName, AudioDir, jsonComp_name+".wav");

            audioTrack = AddAudioTrack(jsonComp_name, trackindex);
            trackindex += 1;
            AddAudioFileToTimeline(fileName, audioTrack, cursorLocation);
        }

        public PosObj Layer_Calculate_Position(VideoTrack vidTrack, Comp jsonComp, PosObj posObj)
        {
            PosObj pos = new PosObj();

            double camRatio = cameraRatio;

            //double fps=CTAScene.fps;
            double transFactorX = sceneRatio.width / sceneRatio.scene;
            double transFactorY = sceneRatio.height / sceneRatio.scene;

            //double xValue = pos.x * transFactorX;
            //double yValue = -pos.z * transFactorY;
            //double zValue = pos.y * camRatio;
            //double keyTime = Math.Round((pos.time - timeShift) * fps * 1000) / 1000;
            //pos.time = Math.Round(keyTime) / fps;

            pos.x = transFactorX * (posObj.x);
            pos.y = transFactorY * (posObj.z + offsety);  // z and y are exchanged
            pos.z = camRatio * posObj.y;

            return pos;

        }

        public RotObj Layer_Calculate_Rotation(RotObj rotobj)
        {
            RotObj calcrot = new RotObj();

            double radian = 57.29577951;

            calcrot.x = rotobj.x * radian;
            calcrot.y = rotobj.z * radian;
            calcrot.z = rotobj.y * radian;

            string s = "Layer_Calculate_Rotation\r\n";
            s += "rotobj.x  " + rotobj.x + "\r\n";
            s += "rotobj.y  " + rotobj.y + "\r\n";
            s += "rotobj.z  " + rotobj.z + "\r\n";

            // show_test_message(s);
            // SaveLogFile(MethodBase.GetCurrentMethod(), s);

            return calcrot;
        }

        public Pivot Layer_Calculate_AnchorPoint(Pivot pivot, PosObj imageCenter)
        {
            Pivot calcanchor = new Pivot();

            double camRatio = this.cameraRatio;
            double sceneRatio = this.sceneRatio.scene;

            double imageCenterFactor = 0; // (double)numParam12.Value / 100; Imagecenter is not needed in VEGAS

            calcanchor.x = imageCenter.x * imageCenterFactor - pivot.x;
            calcanchor.y = imageCenter.y * imageCenterFactor + pivot.y;
            calcanchor.z = imageCenter.z * imageCenterFactor - pivot.z * camRatio;

            return calcanchor;
        }

        public Pivot Layer_Search_Correponding_Pivot(Comp jsonComp, PosObj posObj)
        {
            int time = posObj.time;

            Pivot pivot = new Pivot();
            double sceneRatio = this.sceneRatio.scene;

            foreach (Pivot search_pivot in jsonComp.compObj.pivot)
            {
                if (search_pivot.time <= time)
                {
                    pivot.x = search_pivot.x / sceneRatio;
                    pivot.y = search_pivot.z / sceneRatio;
                    pivot.z = search_pivot.y * this.cameraRatio;

                    return pivot;
                }
            }

            pivot.x = 0;
            pivot.y = 0;
            pivot.z = 0;

            return pivot;
        }

        public ScaleObj Layer_Search_Correponding_Scale(Comp jsonComp, PosObj posObj)
        {
            int time = posObj.time;

            ScaleObj scaleobj = new ScaleObj();
            double sceneRatio = this.sceneRatio.scene;

            foreach (ScaleObj search_scale in jsonComp.compObj.scaleObjs)
            {
                if (search_scale.time <= time)
                {
                    scaleobj.x = search_scale.x;
                    scaleobj.y = search_scale.y;
                    scaleobj.z = search_scale.z;

                    return scaleobj;
                }
            }

            scaleobj.x = 1;
            scaleobj.y = 1;
            scaleobj.z = 1;

            return scaleobj;
        }


        public bool Layer_Add_PositionKey(ref VideoTrack vidTrack, Comp jsonComp)
        {
            CompObj compObj = jsonComp.compObj;
            SaveLogFile(MethodBase.GetCurrentMethod(), "started");
            //var layerScale = layer.scale;
            ScaleObj scale1;
            ScaleObj scale;
            ScaleObj compScale = new ScaleObj();
            compScale.x = jsonComp.compScale.scaleX;
            compScale.y = jsonComp.compScale.scaleZ;
            compScale.z = jsonComp.compScale.scaleY;
            double sceneRatio = this.sceneRatio.scene;

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
                    //tmKF = vidTrack.TrackMotion.InsertMotionKeyframe(kftime);
                }
                PosObj calcpos = Layer_Calculate_Position(vidTrack, jsonComp, posobj);

                Pivot pivot;
                Pivot pivot1;

                
                PosObj imageCenter = new PosObj();
                imageCenter.x = orgWidth / 2;
                imageCenter.y = orgHeight / 2;
                imageCenter.z = 0;
                pivot1 = Layer_Search_Correponding_Pivot(jsonComp, posobj);
                pivot = Layer_Calculate_AnchorPoint(pivot1, imageCenter);
                scale1 = Layer_Search_Correponding_Scale(jsonComp, posobj);
                scale = Layer_Calculate_Scale(scale1, compScale);

                double factorWidth;
                double factorHeight;
               

                PosObj posfactor = new PosObj();
                posfactor.x = 1; // (double)numParam01.Value / 100;
                posfactor.y = 1; // (double)numParam02.Value / 100;
                posfactor.z = 1; // (double)numParam03.Value / 100;

                PosObj pivotfactor = new PosObj();
                pivotfactor.x =  (double)numParam04.Value / 100;
                pivotfactor.y =  (double)numParam05.Value / 100;
                pivotfactor.z = 1; // (double)numParam06.Value / 100;

                if ((orgWidth / orgHeight) > ((double)(SceneWidth) / (double)(SceneHeight))) // orgWidth is the scale factor
                {
                     factorWidth = 1;
                    factorHeight = orgHeight/orgWidth * SceneHeight / SceneWidth;
                }
                else                                       // objheight is the scale factor
                {
                    factorWidth = orgWidth / orgHeight * SceneWidth / SceneHeight;
                    factorHeight = 1;
                }




                tmKF.Type = VideoKeyframeType.Linear;
                //tmKF.PositionX = (posfactor.x * calcpos.x) - (pivotfactor.x * pivot.x) + (pivot.x * scale.x);
                //tmKF.PositionX = calcpos.x - (pivot.x * (1 - scale.x));
                tmKF.PositionX = calcpos.x - (pivot.x * (scale.x)); // * pivotfactor.x * factorWidth);

                //tmKF.PositionY = (posfactor.y * calcpos.y) + (pivotfactor.y * pivot.y) - (pivot.y * scale.y);
                //tmKF.PositionY = calcpos.y + (pivot.y * ( 1 - scale.y));
                tmKF.PositionY = calcpos.y + (pivot.y * (scale.y)); // * pivotfactor.y * factorHeight);
                
                //tmKF.PositionZ = (posfactor.z * calcpos.z) - (pivotfactor.z * pivot.z) + (pivot.z * scale.z);
                //tmKF.PositionZ = calcpos.z - (pivot.z * (1 - scale.z));
                tmKF.PositionZ = calcpos.z + (pivot.z * (scale.z));

                string s = "Layer_Pos_Pivot_Correction\r\n";
                s += jsonComp.name + "\r\n";
                s += posobj.time + "\r\n";
                s += "tmKF-Position" + tmKF.Position.ToPositionString() + "\r\n";
                s += "sceneRatio  " + sceneRatio + "\r\n";
                s += "pivot.x  " + pivot.x + "\r\n";
                s += "pivot.y  " + pivot.y + "\r\n";
                s += "pivot.z  " + pivot.z + "\r\n";
                s += "scale.x  " + scale.x + "\r\n";
                s += "scale.y  " + scale.y + "\r\n";
                s += "scale.z  " + scale.z + "\r\n";
                s += "pivotfactor" + pivotfactor.x + "\r\n";
                s += "pivotfactor" + pivotfactor.y + "\r\n";
                s += "imageCenter.x  " + imageCenter.x + "\r\n";
                s += "imageCenter.y  " + imageCenter.y + "\r\n";
                s += "calcpos.x  " + calcpos.x + "\r\n";
                s += "calcpos.y  " + calcpos.y + "\r\n";
                s += "calcpos.z  " + calcpos.z + "\r\n";
                s += "PositionX  " + tmKF.PositionX + "\r\n";
                s += "PositionY  " + tmKF.PositionY + "\r\n";
                s += "PositionZ  " + tmKF.PositionZ + "\r\n";


                //show_test_message(s);
                SaveLogFile(MethodBase.GetCurrentMethod(), s);

                if (orgWidth > 0)
                {
                    tmKF.Width = orgWidth;
                }
                if (orgHeight > 0)
                {
                    tmKF.Height = orgHeight;
                }
            }

            return true;
        }

        public bool Layer_Add_AnchorPointKey(ref VideoTrack vidTrack, Comp jsonComp)
        {
            return true;

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
        }

        public int lastkey_index = 0;

        public TrackMotionKeyframe Search_Motion_Keyframe(ref VideoTrack vidTrack, double time)
        {
            string s;
            time = scaletime * time;
            if (time == 0)
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
            foreach (TrackMotionKeyframe tmkf in vidTrack.ParentTrackMotion.MotionKeyframes)
            {
                if (tmkf.Position.ToMilliseconds() == time)
                {
                    return tmkf;
                }
            }
            Timecode kftime = Timecode.FromMilliseconds(scaletime * time);
            TrackMotionKeyframe tmkf1 = vidTrack.ParentTrackMotion.InsertMotionKeyframe(kftime);
            //s = "Keyframe not found - new keyframe at: " + tmkf1.Position.ToPositionString();
            //SaveLogFile(MethodBase.GetCurrentMethod(), s);
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
            CompObj compObj = jsonComp.compObj;

            //var layerScale = layer.scale;
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
                    //tmKF = vidTrack.TrackMotion.InsertMotionKeyframe(kftime);
                }
                tmKF.Type = VideoKeyframeType.Linear;

                ScaleObj calcScale = Layer_Calculate_Scale(scaleobj, compScale);

                //double objwidth = orgWidth * calcScale.x;
                //double objheight = orgHeight * calcScale.y;


                if ((orgWidth/orgHeight)>((double)(SceneWidth)/(double)(SceneHeight))) // orgWidth is the scale factor
                {
                    s = "\r\n";
                    s += "Objwidth" + "\r\n";
                    s += "SceneWidth" + SceneWidth + "\r\n";
                    s += "SceneHeight" + SceneHeight + "\r\n";
                    s += "SceneWidth/SceneHeight" + SceneWidth / SceneHeight + "\r\n";
                    //s += "objWidth   " + objwidth + "\r\n";
                    //s += "objHeight  " + objheight + "\r\n";
                    //s += "objwidth/objheight" + objwidth / objheight + "\r\n";
                    SaveLogFile(MethodBase.GetCurrentMethod(), s);

                    tmKF.Width = orgWidth * calcScale.x;
                    tmKF.Height = orgWidth * calcScale.y * SceneHeight / SceneWidth ;
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
                    //s += "objwidth/objheight" + objwidth / objheight + "\r\n";
                    SaveLogFile(MethodBase.GetCurrentMethod(), s);

                    tmKF.Width = orgHeight * calcScale.x * SceneWidth / SceneHeight;
                    tmKF.Height = orgHeight * calcScale.y;
                }

                //if (orgWidth > 0)
                //{
                //    tmKF.Width = orgWidth; //  * calcScale.x
                //}
                //if (orgHeight > 0)
                //{
                //    tmKF.Height = orgHeight * calcScale.y;
                //}
                s = "\r\n";
                s += jsonComp.name + "\r\n";
                s += scaleobj.time + "\r\n";
                s += "tmKF-Position" + tmKF.Position.ToPositionString() + "\r\n";
                s += "orgWidth   " + orgWidth + "\r\n";
                s += "calcScale.x  " + calcScale.x + "\r\n";
                //s += "objWidth   " + objwidth + "\r\n";
                s += "tmKF.Width  " + tmKF.Width + "\r\n";
                s += "sceneRatio.height  " + sceneRatio.height + "\r\n";
                s += "orgHeight  " + orgHeight + "\r\n";
                s += "calcScale.y  " + calcScale.y + "\r\n";
                //s += "objHeight  " + objheight + "\r\n";
                s += "tmKF.Height  " + tmKF.Height + "\r\n";

                // show_test_message(s);
                SaveLogFile(MethodBase.GetCurrentMethod(), s);
            }
            return true;
        }

        public bool Layer_Add_RotationKey(ref VideoTrack vidTrack, Comp jsonComp)
        {
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
                    //tmKF = vidTrack.TrackMotion.InsertMotionKeyframe(kftime);
                }
                tmKF.Type = VideoKeyframeType.Linear;

                RotObj calcRot = Layer_Calculate_Rotation(rotobj);

                tmKF.RotationX = calcRot.x;
                tmKF.RotationY = calcRot.y;
                tmKF.RotationZ = calcRot.z;

            }
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

        public void Layer_Create(Comp jsonComp, ref int trackindex)
        {
            string ImageDir;
            string fileName;
            VideoTrack vidTrack;

            CompObj compObj = jsonComp.compObj;

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
            //Layers_add_Camera("Camera", CTAScene.width / 2, CTAScene.height / 2);

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

            //var cameraPos = aeCamera.position;
            //var cameraPOI = aeCamera.pointOfInterest;
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
                    //tmKF = vidTrack.TrackMotion.InsertMotionKeyframe(kftime);
                }
                PosObj calcpos = Camera_Calculate_Position(posobj);

                PosObj posfactor = new PosObj();
                posfactor.x = -(double)numParam07.Value / 100;
                posfactor.y = (double)numParam08.Value / 100;
                posfactor.z = -0.67 * (double)numParam09.Value / 100;

                PosObj posoffset = new PosObj();
                posoffset.x = SceneWidth * (double)numParam01.Value/100;
                posoffset.y = SceneHeight * (double)numParam02.Value/100;
                posoffset.z = SceneHeight * (double)numParam03.Value/100;
                               
                tmKF.Type = VideoKeyframeType.Linear;
                tmKF.PositionX = posfactor.x * (calcpos.x) + posoffset.x;
                tmKF.PositionY = posfactor.y * (calcpos.y) + posoffset.y;
                tmKF.PositionZ = posfactor.z * (calcpos.z) + posoffset.z;

                //show_test_message(s);
                orgWidth  = (int) ((double) SceneWidth * (double) numParam11.Value /100);
                //orgHeight = (int) ((double) SceneHeight * (double)numParam12.Value / 100);

                if (orgWidth > 0)
                {
                    tmKF.Width = orgWidth;
                    tmKF.Height = orgWidth;
                }

                tmKF.Depth = SceneHeight * (double)numParam10.Value/100;

                string s = "Camera_add_Pos_key\r\n";
                s += "Camera\r\n";
                s += posobj.time + "\r\n";
                s += "calcpos.x  " + calcpos.x + "\r\n";
                s += "calcpos.y  " + calcpos.y + "\r\n";
                s += "calcpos.z  " + calcpos.z + "\r\n";
                //show_test_message(s);
                SaveLogFile(MethodBase.GetCurrentMethod(), s);
            }
            //cameraPos.setValueAtKey(index, [xTrans, yTrans, zTrans]);
            //cameraPOI.setValueAtKey(index, [xTrans, yTrans, zTrans + value]);

            return true;
        }

        public bool Camera_Add_RotateKey(VideoTrack vidTrack, AECamera aeCamera, Camera jsonCamera)
        {

            //var cameraRot = aeCamera.zRotation;

            //var rotKeys = Camera_Calculate_Rotation(jsonCamera);

            //for (var i = 0; i < rotKeys.length; ++i)
            //{

            //    var zRot = rotKeys[i].value;
            //    var keyTime = rotKeys[i].time;

            //    try
            //    {

            //        var index = cameraRot.addKey(keyTime);
            //        cameraRot.setValueAtKey(index, zRot);
            //    }
            //    catch (e)
            //    {

            //alert( e.message );
            //        return false;
            //    }
            //}

            return true;
        }


        public PosObj Camera_Calculate_Position(PosObj posObj)
        {
            PosObj pos = new PosObj();

            double camRatio = cameraRatio;

            //double fps=CTAScene.fps;
            double transFactorX = sceneRatio.width / sceneRatio.scene;
            double transFactorY = sceneRatio.height / sceneRatio.scene;

            //double xValue = pos.x * transFactorX;
            //double yValue = -pos.z * transFactorY;
            //double zValue = pos.y * camRatio;
            //double keyTime = Math.Round((pos.time - timeShift) * fps * 1000) / 1000;
            //pos.time = Math.Round(keyTime) / fps;

            pos.x = transFactorX * (posObj.x);
            pos.y = transFactorY * (-posObj.z);  // z and y are exchanged
            pos.z = camRatio * posObj.y;

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
            SaveLogFile(MethodBase.GetCurrentMethod(), "Started",true);

            string FileName = GetFileName("Please select CartoonAnimator JSON File");

            if (FileName != null)
            {
                Create_Initial();

                jsonString = File.ReadAllText(FileName);
                CTAScene = JsonSerializer.Deserialize<Root>(jsonString);
                maindir = Path.GetDirectoryName(FileName);

                Create_Project();
                Create_Setup();
                Comp_Create();
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

                if (jsonComp.hasAudio==true)
                {
                    Audio_Create(jsonComp.name, ref trackindex);
                }
                Layer_Create(jsonComp, ref trackindex);
            }
        }

        public void Comp_Add_Camera(VideoTrack vidTrack)
        {
            Camera jsonCamera = myCreate.camera;
            //if (isEmptyObject(jsonCamera))
            //{
            //
            //    return;
            //}
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

            // show_test_message(s);
            SaveLogFile(MethodBase.GetCurrentMethod(), s);

            sceneRatio.width = widthRatio;
            sceneRatio.height = heightRatio;
        }

        public double SetCameraRatio()
        {
            var value = CTAScene.Create.camera.camLens * cameraFLM;
            double cameraRatio = (Math.Abs(value / standardZinCTACam));
            double cameraRatio_final = cameraRatio * (double) numParam13.Value / 100;
            
            string s = "cameraRatio   " + cameraRatio + "\r\n";
            s += "CameraRationFactor  " + numParam13.Value + "\r\n";
            s += "cameraRatio_final  " + cameraRatio_final + "\r\n";
            SaveLogFile(MethodBase.GetCurrentMethod(), s);

            return cameraRatio_final;
        }

        public void Camera_Calc_Position(Camera camera, Root CTAScene)
        {
            //double camRatio = cameraRatio;
            var fps = CTAScene.fps;
            var transFactorX = sceneRatio.width / sceneRatio.scene;
            var transFactorY = sceneRatio.height / sceneRatio.scene;
        }

        public string GetFileName(string Title)
        {
            OpenFileDialog svd = new OpenFileDialog();
            svd.Title = Title;
            svd.InitialDirectory = "D:\\CartoonAnimator\\AE_export";
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
                string s = "Count=" + count + "\r\n";
                s += "Filename: " + FileName +"\r\n";
                SaveLogFile(MethodBase.GetCurrentMethod(), s);
            }
            else
            {
                mymedia = myVegas.Project.MediaPool.AddImageSequence(FileName, count, fps);
                string s = "Count=" + count + "\r\n";
                s += "Filename: " + FileName + "\r\n";
                SaveLogFile(MethodBase.GetCurrentMethod(), s);
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
                        int length_ms = (int) (duration);
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
            string location = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) + @"\VEGAS Pro\20.0\";
            //String scriptDirectory = Path.GetDirectoryName(ScriptPortal.Vegas.Script.Directory);
            //String location = scriptDirectory; //Path.Combine(scriptDirectory, "");
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

        private void numParam01_ValueChanged(object sender, EventArgs e)
        {

        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {

        }
    }
}
