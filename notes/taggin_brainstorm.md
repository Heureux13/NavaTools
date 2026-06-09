color = {
    'CLR_Black',
    'CLR_BlackNote',
    'CLR_Blue',
    'CLR_BlueNote',
    'CLR_Cyan',
    'CLR_CyanNote',
    'CLR_Green',
    'CLR_GreenNote',
    'CLR_Magenta',
    'CLR_MagentaNote',
    'CLR_Maroon',
    'CLR_MarronNote
    'CLR_Orange',
    'CLR_OrangeNote',
    'CLR_Purple',
    'CLR_PurpleNote',
    'CLR_Red',
    'CLR_RedNote',
    'CLR_Yellow',
    'CLR_YellowNote',
}

location = {
    '_NV_LOC.Bottom',
    '_NV_LOC.Center',
    '_NV_LOC.Left',
    '_NV_LOC.Right',
    '_NV_LOC.Top',
}

devices = {
    '_Tag.DEV_AFMS',
    '_Tag.DEV_DamperControl',
    '_Tag.DEV_DamperFire'
    '_Tag.DEV_DamperMechanical',
    '_Tag.DEV_DamperVolume',
    '_Tag.DEV_Louver',
    '_Tag.DEV_GRD',
    '_Tag.DEV_AFMS',
}

equipment = {
    '_Tag.EQP_Condenser',
    '_Tag.EQP_Fan',
    '_Tag.EQP_Split',
    '_Tag.EQP_UnitAHU',
    '_Tag.EQP_UnitCRAC',
    '_Tag.EQP_UnitDOAS',
    '_Tag.EQP_VAV',
}

duct = {
    "_NV_DCT.RND_Spiral',
    "_NV_DCT.RND_Reducer',
    "_NV_DCT.RND_Tap',
    "_NV_DCT.SnD_Straight',
    "_NV_DCT.SnD_Reducer',
    "_NV_DCT.TDF_Straight',
    "_NV_DCT.TDF_Straight',
}

bluebeam = {
    '_NV_BBM.Author',
    '_NV_BBM.Ceiling',
    '_NV_BBM.CFMEA',
    '_NV_BBM.CFMSA',
    '_NV_BBM.Class',
    '_NV_BBM.Color',
    '_NV_BBM.Comments',
    '_NV_BBM.Damper',
    '_NV_BBM.Device',
    '_NV_BBM.Duty',
    '_NV_BBM.Face',
    '_NV_BBM.Fan',
    '_NV_BBM.GPM',
    '_NV_BBM.Hand',
    '_NV_BBM.HP',
    '_NV_BBM.K',
    '_NV_BBM.KW',
    '_NV_BBM.Label',
    '_NV_BBM.Layer',
    '_NV_BBM.Make',
    '_NV_BBM.Material',
    '_NV_BBM.Model',
    '_NV_BBM.Mount',
    '_NV_BBM.Neck',
    '_NV_BBM.Notes',
    '_NV_BBM.Number',
    '_NV_BBM.PageLabel',
    '_NV_BBM.Paint',
    '_NV_BBM.Phase',
    '_NV_BBM.Qty',
    '_NV_BBM.Section',
    '_NV_BBM.Size',
    '_NV_BBM.Sleeve',
    '_NV_BBM.Slot',
    '_NV_BBM.Space',
    '_NV_BBM.Status',
    '_NV_BBM.Subclass',
    '_NV_BBM.Subject',
    '_NV_BBM.System',
    '_NV_BBM.Trade',
    '_NV_BBM.Type',
    '_NV_BBM.Unit',
    '_NV_BBM.VAV',
    '_NV_BBM.VPH',
}

python = {
    '_NV_PYT.AspectRatio',
    '_NV_PYT.CFM',
    '_NV_PYT.HeightPad',
    '_NV_PYT.Label',
    '_NV_PYT.Note0',
    '_NV_PYT.Note1',
    '_NV_PYT.Note2',
    '_NV_PYT.Note3',
    '_NV_PYT.Note4',
    '_NV_PYT.Number',
    '_NV_PYT.NumberFabrication',
    '_NV_PYT.NumberRun',
    '_NV_PYT.NumberSleeve',
    '_NV_PYT.OffsetBottom',
    '_NV_PYT.OffsetCenterH',
    '_NV_PYT.OffsetCenterV',
    '_NV_PYT.OffsetLeft',
    '_NV_PYT.OffsetRight',
    '_NV_PYT.OffsetTop',
    '_NV_PYT.OffsetValue',
    '_NV_PYT.SkipNumber',
    '_NV_PYT.SkipTag',
    '_NV_PYT.Sleeve',
    '_NV_PYT.SleeveOpening',
    '_NV_PYT.SleeveValue',
    '_NV_PYT.WeightRun',
    '_NV_PYT.WeightSupport',
}

text_types = {
    '_TXT.125_Arial_Black_R',
    '_TXT.125_Arial_Black.B.',
    '_TXT.125_Arial_Black.U.',
    '_TXT.125_Arial_Black.BU',
}

we start with the trade

Trades = {
'sheetmetal      MDT',
'pipefitting     MPF',
'plumbing        MPL',
}

we add floor

floors = {
    'basement 1 = B01',
    'basement 2 = B02',
    'basement 3 = B03',
    '1st = 001',
    '2nd = 002',
    '3rd = 003',
    '4th = 004',
    '5th = 005',
    '6th = 004',
    '7th = 005',
    '8th = 006',
}

maybe add floor types = {
    'F = floor',
    'B = basement',
    'M = mezzanine',
    'R = roof',
}

we add specifics

specifics = {
    'LVR = louver',
    'SLV = sleeves',
    'SCH = schedules',
    'SPA = supply air',
    'EXA = exhaust air',
    'RTA = return air',
    'OTA = outside air',
    'REA = releaf air',
    'TOA = treated outside air',
    'DTL = details',
    'SCT = section views',
    'DCT = ductwork',
    'DEV = devices',
    'EQP = equipment',
    'EQP = equipment',
    'UNT = unit',
    'FAN = fan',
    'GRD = grilles',
    'DMP = dampers',
    'FDM = fire dampers',
    'MNB = man bars',
    'VAV = vav',
    'VAV = vav',
}

we add area

Areas = {
    'Area A = A'
    'Area B = B'
    'Area C = C'
    'Area D = D'
    'Area E = E'
}

we add area number

area number = {
    'A01',
    'A02',
    'A03',
    'A04',
    'B01',
    'B02',
    'B03',
    'B04',
}

UIM-MDT-DCT-01-A1
MDT-SLV-01-A1

MDT-F01-A01
MDT-DTL-01
MDT-SLV-A01
MDT-SLV-DT1


MDT-F01-A1
MDT-F01-A2
MDT-F01-A3
MDT-F01-A4
MDT-F01-B1
MDT-F01-B2
MDT-F01-B3
MDT-F01-B4

