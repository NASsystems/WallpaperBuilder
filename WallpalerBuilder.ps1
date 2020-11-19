<#
.SYNOPSIS
壁紙ビルダー
.DESCRIPTION
壁紙ビルダー
.NOTES
MIT License

Copyright (c) 2020 Isao Sato

Permission is hereby granted, free of charge, to any person obtaining a
copy of this software and associated documentation files (the
"Software"), to deal in the Software without restriction, including
without limitation the rights to use, copy, modify, merge, publish,
distribute, sublicense, and/or sell copies of the Software, and to
permit persons to whom the Software is furnished to do so, subject to
the following conditions:

The above copyright notice and this permission notice shall be included
in all copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS
OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF
MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT.
IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY
CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT,
TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE
SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
#>

[System.Reflection.Assembly]::LoadWithPartialName('System.Windows.Forms') | Out-Null

# file format
$mimetype = "image/jpeg"
$encparams = New-Object System.Drawing.Imaging.EncoderParameters -ArgumentList 1
$encparams.Param[0] = New-Object System.Drawing.Imaging.EncoderParameter -ArgumentList @([System.Drawing.Imaging.Encoder]::Quality, [System.Int64] 80)


function Win32 {
	$DefiningTypes = @{}
	$Win32Type = New-Object System.Collections.Generic.Dictionary[string`,type]
	
	$appdomain = [AppDomain]::CurrentDomain
	$asmbuilder = $appdomain.DefineDynamicAssembly((New-Object Reflection.AssemblyName 'Win32'), [Reflection.Emit.AssemblyBuilderAccess]::Run)
	$modbuilder = $asmbuilder.DefineDynamicModule('Win32.dll')
	
	$modbuilder |% {
		$_.DefineType(
			'WallpaperHelper.Win32.User32',
			[System.Reflection.TypeAttributes] 'AutoLayout, AnsiClass, Class, Public, BeforeFieldInit',
			[System.Object]
			) |% {
			$DefiningTypes['WallpaperHelper.Win32.User32'] = @{}
			$DefiningTypes['WallpaperHelper.Win32.User32'].Builder = $_
			$_.DefineNestedType(
				'SPI',
				[System.Reflection.TypeAttributes] 'AutoLayout, AnsiClass, Class, NestedPublic, Sealed',
				[System.Enum]
				) |% {
				$DefiningTypes['WallpaperHelper.Win32.User32+SPI'] = @{}
				$DefiningTypes['WallpaperHelper.Win32.User32+SPI'].Builder = $_
			} | Out-Null
			$_.DefineNestedType(
				'SPIF',
				[System.Reflection.TypeAttributes] 'AutoLayout, AnsiClass, Class, NestedPublic, Sealed',
				[System.Enum]
				) |% {
				$DefiningTypes['WallpaperHelper.Win32.User32+SPIF'] = @{}
				$DefiningTypes['WallpaperHelper.Win32.User32+SPIF'].Builder = $_
			} | Out-Null
		} | Out-Null
	}
		
	function Create-CustomAttributeBuilder([Reflection.ConstructorInfo] $constructor, [object[]] $arguments, [System.Collections.Hashtable] $attributes)
	{
		$AttributeFields = New-Object Collections.Generic.List[Reflection.FieldInfo]
		$AttributeValues = New-Object Collections.Generic.List[Object]
			
		$attributes.GetEnumerator() |% {
			$AttributeFields.Add($constructor.ReflectedType.GetField($_.Key)) | Out-Null
			$AttributeValues.Add($_.Value) | Out-Null
		}
			
		New-Object Reflection.Emit.CustomAttributeBuilder ($constructor, $arguments, $AttributeFields.ToArray(), $AttributeValues.ToArray())
	}
		
	$DefiningTypes['WallpaperHelper.Win32.User32'].Builder |% {
		$_ | Out-Null
		$_.SetCustomAttribute(
			(Create-CustomAttributeBuilder `
				([System.Runtime.InteropServices.StructLayoutAttribute].GetConstructor(@([System.Runtime.InteropServices.LayoutKind]))) `
				@(([System.Runtime.InteropServices.LayoutKind] 'Auto')) `
				@{
					CharSet = ([System.Runtime.InteropServices.CharSet] 'Ansi')
					Pack = 8
					Size = 0
				}
				)
			) | Out-Null
		$_.DefineMethod(
			'SystemParametersInfo',
			[System.Reflection.MethodAttributes] 'PrivateScope, Public, Static, HideBySig, PinvokeImpl',
			[System.Boolean],
			@($DefiningTypes['WallpaperHelper.Win32.User32+SPI'].Builder, [System.Int32], [System.Text.StringBuilder], $DefiningTypes['WallpaperHelper.Win32.User32+SPIF'].Builder)
			) |% {
			$_.SetCustomAttribute(
				(Create-CustomAttributeBuilder `
					([System.Runtime.InteropServices.DllImportAttribute].GetConstructor(@([System.String]))) `
					@(([System.String] 'user32.dll')) `
					@{
						EntryPoint = 'SystemParametersInfo' # [System.String]
						CharSet = 1 # [System.Runtime.InteropServices.CharSet]
						SetLastError = $true # [System.Boolean]
						PreserveSig = $true # [System.Boolean]
						CallingConvention = 1 # [System.Runtime.InteropServices.CallingConvention]
					}
					)
				) | Out-Null
			$_.SetCustomAttribute(
				(Create-CustomAttributeBuilder `
					([System.Runtime.InteropServices.PreserveSigAttribute].GetConstructor(@())) `
					@() `
					@{
					}
					)
				) | Out-Null
			$_.SetImplementationFlags([System.Reflection.MethodImplAttributes] 'PreserveSig') | Out-Null
			$_.DefineParameter(
				1,
				[System.Reflection.ParameterAttributes] 'None',
				'uAction'
				) | Out-Null
			$_.DefineParameter(
				2,
				[System.Reflection.ParameterAttributes] 'None',
				'uParam'
				) | Out-Null
			$_.DefineParameter(
				3,
				[System.Reflection.ParameterAttributes] 'None',
				'lpvParam'
				) | Out-Null
			$_.DefineParameter(
				4,
				[System.Reflection.ParameterAttributes] 'None',
				'fuWinIini'
				) | Out-Null
		} | Out-Null
	} | Out-Null
	$DefiningTypes['WallpaperHelper.Win32.User32+SPI'].Builder |% {
		$_.SetCustomAttribute(
			(Create-CustomAttributeBuilder `
				([System.Runtime.InteropServices.StructLayoutAttribute].GetConstructor(@([System.Runtime.InteropServices.LayoutKind]))) `
				@(([System.Runtime.InteropServices.LayoutKind] 'Auto')) `
				@{
					CharSet = ([System.Runtime.InteropServices.CharSet] 'Ansi')
					Pack = 8
					Size = 0
				}
				)
			) | Out-Null
		$_.DefineField(
			'value__',
			[System.Int32],
			[System.Reflection.FieldAttributes] 'Public, SpecialName, RTSpecialName'
			) | Out-Null
		$_.DefineField(
			'SETDESKWALLPAPER',
			$DefiningTypes['WallpaperHelper.Win32.User32+SPI'].Builder,
			[System.Reflection.FieldAttributes] 'Public, Static, Literal, HasDefault'
			) |% {
			$_.SetConstant(20)
		} | Out-Null
		$_.DefineField(
			'GETDESKWALLPAPER',
			$DefiningTypes['WallpaperHelper.Win32.User32+SPI'].Builder,
			[System.Reflection.FieldAttributes] 'Public, Static, Literal, HasDefault'
			) |% {
			$_.SetConstant(115)
		} | Out-Null
	} | Out-Null
	$DefiningTypes['WallpaperHelper.Win32.User32+SPIF'].Builder |% {
		$_.SetCustomAttribute(
			(Create-CustomAttributeBuilder `
				([System.Runtime.InteropServices.StructLayoutAttribute].GetConstructor(@([System.Runtime.InteropServices.LayoutKind]))) `
				@(([System.Runtime.InteropServices.LayoutKind] 'Auto')) `
				@{
					CharSet = ([System.Runtime.InteropServices.CharSet] 'Ansi')
					Pack = 8
					Size = 0
				}
				)
			) | Out-Null
		$_.DefineField(
			'value__',
			[System.Int32],
			[System.Reflection.FieldAttributes] 'Public, SpecialName, RTSpecialName'
			) | Out-Null
		$_.DefineField(
			'SPIF_NONE',
			$DefiningTypes['WallpaperHelper.Win32.User32+SPIF'].Builder,
			[System.Reflection.FieldAttributes] 'Public, Static, Literal, HasDefault'
			) |% {
			$_.SetConstant(0)
		} | Out-Null
		$_.DefineField(
			'SPIF_UPDATEINIFILE',
			$DefiningTypes['WallpaperHelper.Win32.User32+SPIF'].Builder,
			[System.Reflection.FieldAttributes] 'Public, Static, Literal, HasDefault'
			) |% {
			$_.SetConstant(1)
		} | Out-Null
		$_.DefineField(
			'SPIF_SENDCHANGE',
			$DefiningTypes['WallpaperHelper.Win32.User32+SPIF'].Builder,
			[System.Reflection.FieldAttributes] 'Public, Static, Literal, HasDefault'
			) |% {
			$_.SetConstant(2)
		} | Out-Null
	} | Out-Null
    
	@(
		'WallpaperHelper.Win32.User32',
		'WallpaperHelper.Win32.User32+SPI',
		'WallpaperHelper.Win32.User32+SPIF'
	) |% {$Win32Type[$_] = $DefiningTypes[$_].Builder.CreateType()}

	$Win32Type.GetEnumerator() |% {
		[psobject].Assembly.GetType('System.Management.Automation.TypeAccelerators')::Add(
			$_.Key,
			$_.Value
			)
	}

	$Win32Type.GetEnumerator() |% {
		[psobject].Assembly.GetType('System.Management.Automation.TypeAccelerators')::Add(
			$_.Key+'[]',
			$_.Value.MakeArrayType()
			)
	}
}


function DrawText {
    param(
        [System.Drawing.Graphics] $Graphics,
        [string] $Text,
        [System.Drawing.RectangleF] $TextArea,
        [System.Drawing.Font] $Font,
        [System.Drawing.StringFormat] $StringFormat,
        [System.Drawing.Brush] $Brush,
        [System.Drawing.Pen] $Outline,
        [System.Drawing.Drawing2D.Matrix] $Matrix = (New-Object System.Drawing.Drawing2D.Matrix))
    
    #$p = New-Object System.Drawing.Pen ([System.Drawing.Brushes]::Blue), 1
    #$Graphics.DrawRectangle($p, ([int]$TextArea.Left), ([int]$TextArea.Top), ([int]$TextArea.Width), ([int]$TextArea.Height))
    #$p.Dispose()
    
    $path = New-Object System.Drawing.Drawing2D.GraphicsPath
    $path.AddString(
        $Text,
        $Font.Name,
        [int] $Font.Style,
        $Font.Size,
        $TextArea,
        $StringFormat)
    $path.Transform($Matrix)
    $Graphics.DrawPath($Outline, $path)
    $Graphics.FillPath($Brush, $path)
    
    $rect = $path.GetBounds()
    
    $path.Dispose()
    
    $rect
}


function MeasureText {
    param(
        [System.Drawing.Graphics] $Graphics,
        [string] $Text,
        [System.Drawing.RectangleF] $TextArea,
        [System.Drawing.Font] $Font,
        [System.Drawing.StringFormat] $StringFormat)
    
    $path = New-Object System.Drawing.Drawing2D.GraphicsPath
    $path.AddString(
        $Text,
        $Font.Name,
        [int] $Font.Style,
        $Font.Size,
        $TextArea,
        $StringFormat)
    #$Graphics.DrawPath($Outline, $path)
    #$Graphics.FillPath($Brush, $path)
    
    $rect = $path.GetBounds()
    
    $path.Dispose()
    
    $rect
}

Win32

$lpvParam = New-Object System.Text.StringBuilder 260
[WallpaperHelper.Win32.User32]::SystemParametersInfo(
    ([WallpaperHelper.Win32.User32+SPI]::GETDESKWALLPAPER),
    $lpvParam.Capacity,
    $lpvParam,
    ([WallpaperHelper.Win32.User32+SPIF]::SPIF_NONE))
$sourcebitmappath = $lpvParam.ToString()

$bmp0 = $null
try {
    $bmp0 = New-Object System.Drawing.Bitmap $sourcebitmappath
    $bmp1 = New-Object System.Drawing.Bitmap $bmp0
} finally {
    if($bmp0){$bmp0.Dispose()}
}

$g = [System.Drawing.Graphics]::FromImage($bmp1)
$g.TextRenderingHint = [System.Drawing.Text.TextRenderingHint]::AntiAlias

$screensize = [System.Windows.Forms.Screen]::PrimaryScreen.Bounds.Size
$bmpsize = $bmp1.Size
if(($bmpsize.Width / $bmpsize.Height) -ge ($screensize.Width / $screensize.Height)) {
    $screenbox = New-Object System.Drawing.Rectangle -ArgumentList @(
        (New-Object System.Drawing.Point -ArgumentList (($bmpsize.Width -($screensize.Width * $bmpsize.Height / $screensize.Height))/2), 0), 
        (New-Object System.Drawing.Size  -ArgumentList (($screensize.Width * $bmpsize.Height / $screensize.Height), $bmpsize.Height)))
} else {
    $screenbox = New-Object System.Drawing.Rectangle -ArgumentList @(
        (New-Object System.Drawing.Point 0, (($bmpsize.Height -($screensize.Height * $bmpsize.Width / $screensize.Width))/2)), 
        (New-Object System.Drawing.Size  $bmpsize.Width, ($screensize.Height * $bmpsize.Width / $screensize.Width)))
}


$fontfamily = New-Object System.Drawing.FontFamily 'メイリオ'
$text1 = ''

$text1area = New-Object System.Drawing.RectangleF -ArgumentList @(
    (New-Object System.Drawing.PointF ($screenbox.Left +$screenbox.Width /16),  ($screenbox.Top +$screenbox.Height / 10)), 
    (New-Object System.Drawing.SizeF  ($screenbox.Width -$screenbox.Width /16 *2), ($screenbox.Height / 10)))

$text1size   = $screenbox.Height / 10
$text1font = New-Object System.Drawing.Font $fontfamily, $text1size
$text1format = New-Object System.Drawing.StringFormat ([System.Drawing.StringFormat]::GenericDefault)
$text1format.Alignment = [System.Drawing.StringAlignment]::Far

$text1brush = [System.Drawing.Brushes]::White
$text1outline = New-Object System.Drawing.Pen ([System.Drawing.Color]::FromArgb(128, 0, 0, 0)), (($screenbox.Width +$screenbox.Height) / 700)

DrawText $g $text1 $text1area $text1font $text1format $text1brush $text1outline | Out-Null

$text1font.Dispose()

$text2 = ''
$text2area = New-Object System.Drawing.RectangleF -ArgumentList @(
    ($text1area.Location +(New-Object System.Drawing.SizeF 0, ($text1area.Height *1.8))), 
    $text1area.Size)
$text2size   = $text1size *0.5
$text2font = New-Object System.Drawing.Font $fontfamily, $text2size

$text2format = New-Object System.Drawing.StringFormat ([System.Drawing.StringFormat]::GenericDefault)
$text2format.Alignment = [System.Drawing.StringAlignment]::Far

$text2brush = [System.Drawing.Brushes]::White
$text2outline = New-Object System.Drawing.Pen ([System.Drawing.Color]::FromArgb(128, 0, 0, 0)), (($screenbox.Width +$screenbox.Height) / 700)

DrawText $g $text2 $text2area $text2font $text2format $text2brush $text2outline | Out-Null

$text2font.Dispose()

$text3a1 = 'ホスト名: '
$text3a2 = 'ユーザ名: '
$text3b1 = $env:COMPUTERNAME
$text3b2 = '{0}\{1}' -f $env:USERDOMAIN, $env:USERNAME

$text3size   = $text2size

$text3font = New-Object System.Drawing.Font $fontfamily, $text3size
$text3format = New-Object System.Drawing.StringFormat ([System.Drawing.StringFormat]::GenericDefault)
$text3area = New-Object System.Drawing.RectangleF -ArgumentList @(
    ($text2area.Location +(New-Object System.Drawing.SizeF 0, ($text1area.Height *1.8))), 
    $text1area.Size)
$text3a1size = (MeasureText $g $text3a1 $text3area $text3font $text3format).Size
$text3a2size = (MeasureText $g $text3a2 $text3area $text3font $text3format).Size
$text3b1size = (MeasureText $g $text3b1 $text3area $text3font $text3format).Size
$text3b2size = (MeasureText $g $text3b2 $text3area $text3font $text3format).Size

$text3awidth = [Math]::Max($text3a1size.Width, $text3a2size.Width) +$text3size/3
$text3bwidth = [Math]::Max($text3b1size.Width, $text3b2size.Width)
$text31height = [Math]::Max($text3a1size.Height, $text3b1size.height)
$text32height = [Math]::Max($text3a2size.Height, $text3b2size.height)

$text3a1matrix = New-Object System.Drawing.Drawing2D.Matrix
$text3a1matrix.Translate($text3area.Width -($text3bwidth +$text3awidth) -$text3size, 0)

$text3a2matrix = New-Object System.Drawing.Drawing2D.Matrix
$text3a2matrix.Translate($text3area.Width -($text3bwidth +$text3awidth) -$text3size, $text31height *2)

$text3b1matrix = New-Object System.Drawing.Drawing2D.Matrix
$text3b1matrix.Translate($text3area.Width -$text3bwidth -$text3size, 0)

$text3b2matrix = New-Object System.Drawing.Drawing2D.Matrix
$text3b2matrix.Translate($text3area.Width -$text3bwidth -$text3size, $text31height *2)

$text3abrush = [System.Drawing.Brushes]::White
$text3aoutline = New-Object System.Drawing.Pen ([System.Drawing.Color]::FromArgb(30, 0, 0, 0)), (($screenbox.Width +$screenbox.Height) / 700)

$text3bbrush = [System.Drawing.Brushes]::OrangeRed
$text3boutline = New-Object System.Drawing.Pen ([System.Drawing.Color]::FromArgb(30, 0, 0, 0)), (($screenbox.Width +$screenbox.Height) / 700)

DrawText $g $text3a1 $text3area $text3font $text3format $text3abrush $text3aoutline $text3a1matrix | Out-Null
DrawText $g $text3a2 $text3area $text3font $text3format $text3abrush $text3aoutline $text3a2matrix | Out-Null
DrawText $g $text3b1 $text3area $text3font $text3format $text3bbrush $text3boutline $text3b1matrix | Out-Null
DrawText $g $text3b2 $text3area $text3font $text3format $text3bbrush $text3boutline $text3b2matrix | Out-Null

$text3font.Dispose()
$fontfamily.Dispose()

$g.Dispose()

$codecinfo = [System.Drawing.Imaging.ImageCodecInfo]::GetImageEncoders() | Where-Object {$_.MimeType -eq $mimetype} | Select-Object –First 1
$pictext = [System.IO.Path]::GetExtension($codecinfo.FilenameExtension.Split(';')[0])

$PicturePath = Join-Path ([System.Environment]::GetFolderPath([System.Environment+SpecialFolder]::MyPictures)) ([System.IO.Path]::ChangeExtension('wallpaler.pct', $pictext))

$bmp1.Save(
    $PicturePath,
    $codecinfo,
    $encparams)
$bmp1.Dispose()

$lpvParam = New-Object System.Text.StringBuilder $PicturePath
[WallpaperHelper.Win32.User32]::SystemParametersInfo(
    ([WallpaperHelper.Win32.User32+SPI]::SETDESKWALLPAPER),
    0,
    $lpvParam,
    ([WallpaperHelper.Win32.User32+SPIF]::SPIF_SENDCHANGE))
