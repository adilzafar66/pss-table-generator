ANSI_INT_QUERY = (r"SELECT PDID, kVnom, FaultedBus, PDType, AdjSym, CapAdjInt "
                  r"FROM SCDSumInt WHERE TRIM(PDID) <> '' ORDER BY PDID ASC")
ANSI_MOM_QUERY = (r"SELECT PDID, kVnom, PDType, kASymm, kAASymm, CapSym, CapAsym "
                  r"FROM SCDSumMom WHERE TRIM(PDID) <> '' ORDER BY PDID ASC")
ANSI_INT_SP_QUERY = (r"SELECT PDID, kVnom, FaultedBus, PDType, AdjSym, CapAdjInt "
                     r"FROM SCDSumInt1Ph WHERE TRIM(PDID) <> '' ORDER BY PDID ASC")
ANSI_MOM_SP_QUERY = (r"SELECT PDID, kVnom, PDType, kASymm, kAASymm, CapSym, CapAsym "
                     r"FROM SCDSumMom1Ph WHERE TRIM(PDID) <> '' ORDER BY PDID ASC")
IEC_INT_QUERY = (r"SELECT DeviceID, kVnom, FaultedBus, DeviceType, Ibsymm, Ibasymm, DeviceIbsymm, "
                 r"DeviceIbasym FROM SCIEC3phSum WHERE TRIM(DeviceID) <> '' ORDER BY DeviceID ASC")
IEC_INT_SP_QUERY = (r"SELECT DeviceID, kVnom, FaultedBus, DeviceType, Ibsymm, Ibasymm, DeviceIbsymm, "
                    r"DeviceIbasym FROM SCIEC1phSum WHERE TRIM(DeviceID) <> '' ORDER BY DeviceID ASC")

ANSI_AF_INFO_QUERY = "SELECT Output, Config FROM IAFStudyCase"
ANSI_AF_BUS_QUERY = (
    "SELECT IDBus, NomlkV, EqType, Orientation, WDistance, FixedBoundary, ResBoundary, "
    "IEnergy, PBoundary, FCT, FCTPD, ArcVaria, ArcI, FCTPDIa, FaultI, FCTPDIf "
    "FROM BusArcFlash WHERE EqType <> 'Cable Bus'"
)
ANSI_AF_PD_QUERY = (
    "SELECT ID, NomlkV, Type, Orientation, WDistance, FixedBoundary, ResBoundary, "
    "IEnergy, PBoundary, EnFCT, FCTPD, ArcVaria, EnIa, FCTPDIa, EnIf, FCTPDIf "
    "FROM PDArcFlash WHERE ID <> '' AND Type = 'SPST Switch'"
)
