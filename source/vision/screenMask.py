#vision/screenMask.py
#A part of NonVisual Desktop Access (NVDA)
#This file is covered by the GNU General Public License.
#See the file COPYING for more details.
#Copyright (C) 2018 NV Access Limited

from . import ColorEnhancer, ColorTransformation
import winMagnification
from ctypes import byref
import winVersion

TRANSFORM_BLACK = winMagnification.MAGCOLOREFFECT()
TRANSFORM_BLACK.transform[4][4] = 1.0
TRANSFORM_DEFAULT = winMagnification.MAGCOLOREFFECT()
TRANSFORM_DEFAULT.transform[0][0] = 1.0
TRANSFORM_DEFAULT.transform[1][1] = 1.0
TRANSFORM_DEFAULT.transform[2][2] = 1.0
TRANSFORM_DEFAULT.transform[3][3] = 1.0
TRANSFORM_DEFAULT.transform[4][4] = 1.0

class WinMagnificationScreenMask(ColorEnhancer):
	"""Screen mask implementation based on the windows magnification API.
	This is only supported on Windows 8 and abbove."""
	name = "screenMask"
	availableTransformations = (
		ColorTransformation("black", _("black screen"), TRANSFORM_BLACK)
	)

	def __init__(self, *roles):
		if (winVersion.major, winVersion.minor) < (6, 2):
			raise RuntimeError("This vision enhancement provider is only supported on Windows 8 and above")
		winMagnification.Initialize()
		super(ScreenMask, self).__init__(*roles)

	def initializeColorEnhancer(self):
		super(ScreenMask, self).initializeColorEnhancer()
		winMagnification.SetFullscreenColorEffect(byref(TRANSFORM_BLACK))

	def terminateColorEnhancer(self):
		winMagnification.SetFullscreenColorEffect(byref(TRANSFORM_DEFAULT))
		super(ScreenMask, self).terminateColorEnhancer()
