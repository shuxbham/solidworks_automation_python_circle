#!/usr/bin/env python
# coding: utf-8

# In[30]:


import win32com.client
import pythoncom


# In[31]:


swApp = win32com.client.Dispatch("SldWorks.Application")
swApp.Visible=True


# In[32]:


template=r"C:\ProgramData\SolidWorks\SOLIDWORKS 2024\templates\Part.prtdot"


# In[33]:


Part=swApp.NewDocument(template,0,0,0)


# In[27]:


boolstatus = Part.Extension.SelectByID2("Front Plane", "PLANE", 0, 0, 0, False, 0,pythoncom.Nothing, 0)
Part.SketchManager.InsertSketch(True)
Part.ClearSelection2(True)


# In[34]:


skSegment = Part.SketchManager.CreateCircle(0, 0, 0, 0.05, 0, 0)
Part.ClearSelection2(True)
Part.SketchManager.InsertSketch(True)


# In[35]:


Part.ClearSelection2(True)

