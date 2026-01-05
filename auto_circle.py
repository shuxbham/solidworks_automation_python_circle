{
 "cells": [
  {
   "cell_type": "code",
   "execution_count": 30,
   "id": "dbda5c17-a23f-45fb-8311-76c6a575faef",
   "metadata": {},
   "outputs": [],
   "source": [
    "import win32com.client\n",
    "import pythoncom"
   ]
  },
  {
   "cell_type": "code",
   "execution_count": 31,
   "id": "0aceecc8-6760-4109-8236-320f43a1b9b1",
   "metadata": {},
   "outputs": [],
   "source": [
    "swApp = win32com.client.Dispatch(\"SldWorks.Application\")\n",
    "swApp.Visible=True\n"
   ]
  },
  {
   "cell_type": "code",
   "execution_count": 32,
   "id": "c49527a3-9785-4b91-a091-5db283e8d469",
   "metadata": {},
   "outputs": [],
   "source": [
    "template=r\"C:\\ProgramData\\SolidWorks\\SOLIDWORKS 2024\\templates\\Part.prtdot\""
   ]
  },
  {
   "cell_type": "code",
   "execution_count": 33,
   "id": "e99c5a4c-d666-4241-83c5-03c61670f5d2",
   "metadata": {},
   "outputs": [],
   "source": [
    "Part=swApp.NewDocument(template,0,0,0)"
   ]
  },
  {
   "cell_type": "code",
   "execution_count": 27,
   "id": "77e38547-be42-4da2-979b-b1f3025f9a76",
   "metadata": {},
   "outputs": [],
   "source": [
    "boolstatus = Part.Extension.SelectByID2(\"Front Plane\", \"PLANE\", 0, 0, 0, False, 0,pythoncom.Nothing, 0)\n",
    "Part.SketchManager.InsertSketch(True)\n",
    "Part.ClearSelection2(True)"
   ]
  },
  {
   "cell_type": "code",
   "execution_count": 34,
   "id": "b054b1e0-9676-4a22-ac1f-48a8ab8253d3",
   "metadata": {},
   "outputs": [],
   "source": [
    "skSegment = Part.SketchManager.CreateCircle(0, 0, 0, 0.05, 0, 0)\n",
    "Part.ClearSelection2(True)\n",
    "Part.SketchManager.InsertSketch(True)"
   ]
  },
  {
   "cell_type": "code",
   "execution_count": 35,
   "id": "7116f379-1ee9-41dc-9e2d-35d5a6cc13aa",
   "metadata": {},
   "outputs": [],
   "source": [
    "Part.ClearSelection2(True)"
   ]
  }
 ],
 "metadata": {
  "kernelspec": {
   "display_name": "Python 3 (ipykernel)",
   "language": "python",
   "name": "python3"
  },
  "language_info": {
   "codemirror_mode": {
    "name": "ipython",
    "version": 3
   },
   "file_extension": ".py",
   "mimetype": "text/x-python",
   "name": "python",
   "nbconvert_exporter": "python",
   "pygments_lexer": "ipython3",
   "version": "3.13.7"
  }
 },
 "nbformat": 4,
 "nbformat_minor": 5
}
