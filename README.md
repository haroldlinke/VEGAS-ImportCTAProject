# VEGAS-ImportCTAProject


=======
This project provides a custom command extension for VEGAS Pro that allows the import of CartoonAnimator projects.

# Installation:
- Download the zip archive Vegas-CTAImport.zip from the last release in this github repository.

- Unzip the content into the folder Documents\Vegas Application Extensions.

- After restart of VEGAS Pro you can activate the extension by selecting the option
"view->extensions->Import CTA projects" a dockable window will open. You can dock this window in the VEGAS Pro UI where you want or keep it floating.

# Usage
The extension provides one Button: "Import CTA Project".

By clicking this button you can select a JSON-file created by CartoonAnimator.

You create this JSON-file by selection the render option "export to After Effects" in CartoonAnimator.

By importing the CartoonAnimator JSON-file all characters and props including their movements are imported to VEGAS tracks. Each character and each prop are loaded an individual tracks.
The movements and positions of the objects in the CartoonAnimator 3D-space are translated to 3D-trackmotion keyframes of VEGAS.

The CartoonAnimator camera is simulated by adding a parent track called "Camera" to all imported CartoonAnimator tracks (making them child tracks). 
The camera movement in CartoonAnimator is simulated by corresponding 3D-motion keyframes of the parent track.
Unfortunately this simulation is not perfect. I have therefore added the possibility to finetune the camera movement.

In the "finetune camera keyframes" you can add offsets in pixels to the x,y and z coordinates of the camera track, as well as a stretch factor in percent to the x, y and z calculation.
By clicking the button "Recalc Camera Keyframes" the changes you made to the values are adapted to all keyframes of the camera track.

Attention: The values are not saved!!

The imported tracks are added before the first track at the current cursor position. This allows you to add a CartoonAnimator scene to an existing project.

You can now add further video tracks and position them at any layer in the 3-D space by changing the z-position in 3D-trackmotion. 

If you want the added video to follow the camera movements, make it a child to the “camera” track.

Make sure that the export parameters of CartoonAnimator (size, width and height and fps) fit with the project settings in VEGAS.
There is no check if the parameters are compatible. It might be on purpose.

# known issues
- trying to import a not compatible JSON-File might cause a crash to the command with a cryptic error message coming from VEGAS that something was going wrong.
- As the 3-D models of Cartoon-Animator and VEGAS are completely different the resulting scene may not always look identical in CartoonAnimator and VEGAS. I tried to make it look as similar as possible.
- If you find visible errors, it would be nice to inform me and provide me with the CartoonAnimator project so that I can reproduce the issue.
