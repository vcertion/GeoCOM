Attribute VB_Name = "STUBS32P"






'/*****************************************************************************/
'/*                                                                           */
'/* Project:                 GEOCOM                                           */
'/*                                                                           */
'/* Module Name:             STUBS32.BAS                                      */
'/*                                                                           */
'/* Description:             Generated remote-procedure-call interface module */
'/*                          This file should not need to be edited.          */
'/*                          If there are errors contact Greg Ennis@Leica ATL */
'/*                                                                           */
'/* RPC-Generator version:   2.1                                              */
'/* RPC-Generator date:      Apr  2 1997                                      */
'/* File Generation date:    Wed Apr 23 08:36:14 1997                         */
'/*                                                                           */
'/* Development Environment: Microsoft Visual Basic 4.0                       */
'/*                                                                           */
'/*****************************************************************************/

#If Win32 Then

' Global constant declarations
Global Const PI = 3.1415926
Global Const MAX_APPLICATIONS = 15
Global Const APP_ID_SYSTEM = 1000
Global Const APP_ID_ILLEGAL = -1
Global Const NO_OF_LANGUAGES = 3
Global Const MAX_GENERIC_LENGTH = 450
Global Const BASE_GROUPS = 11
Global Const MAX_USERS = 5
Global Const MIN_MEMACC_VOLTAGE = 11.8
Global Const MAX_PERCENT = 100
Global Const RCBETA = 0
Global Const RC_ALT = 8448
Global Const RC_ANG = 256
Global Const RC_ATA = 512
Global Const RC_AUS = 8960
Global Const RC_AUT = 8704
Global Const RC_BAP = 9216
Global Const RC_BAS = 9984
Global Const RC_BMM = 2304
Global Const RC_COD = 9728
Global Const RC_COM = 3072
Global Const RC_CSV = 4096
Global Const RC_CTL = 4352
Global Const RC_DBM = 3328
Global Const RC_DEL = 3584
Global Const RC_DPL = 4864
Global Const RC_EDM = 768
Global Const RC_FIL = 3840
Global Const RC_GMF = 1024
Global Const RC_LDR = 2048
Global Const RC_MEM = 1536
Global Const RC_MMI = 2816
Global Const RC_MOT = 1792
Global Const RC_SAP = 9472
Global Const RC_STP = 4608
Global Const RC_SUP = 0
Global Const RC_TMC = 1280
Global Const RC_TXT = 2560
Global Const RC_USR = 8192
Global Const RC_WIR = 5120
Global Const RC_APP = 20480
Global Const RC_RES = 28672
Global Const RC_OK = 0
Global Const RC_UNDEFINED = 1
Global Const RC_IVPARAM = 2
Global Const RC_IVRESULT = 3
Global Const RC_FATAL = 4
Global Const RC_NOT_IMPL = 5
Global Const RC_TIME_OUT = 6
Global Const RC_SET_INCOMPL = 7
Global Const RC_ABORT = 8
Global Const RC_NOMEMORY = 9
Global Const RC_NOTINIT = 10
Global Const RC_SHUT_DOWN = 12
Global Const RC_SYSBUSY = 13
Global Const RC_HWFAILURE = 14
Global Const RC_ABORT_APPL = 15
Global Const RC_LOW_POWER = 16
Global Const BMM_NO_MMI_MOD = 0
Global Const BMM_MMI_MOD1 = 1
Global Const BMM_MMI_MOD2 = 2
Global Const BMM_MMI_MOD1_2 = 3
Global Const BMM_NUM0_KEY = 528
Global Const BMM_NUM1_KEY = 544
Global Const BMM_NUM2_KEY = 1056
Global Const BMM_NUM3_KEY = 2080
Global Const BMM_NUM4_KEY = 576
Global Const BMM_NUM5_KEY = 1088
Global Const BMM_NUM6_KEY = 2112
Global Const BMM_NUM7_KEY = 640
Global Const BMM_NUM8_KEY = 1152
Global Const BMM_NUM9_KEY = 2176
Global Const BMM_SIGN_KEY = 2064
Global Const BMM_POINT_KEY = 1040
Global Const BMM_ESC_KEY = 260
Global Const BMM_CE_KEY = 520
Global Const BMM_SHIFT_KEY = 264
Global Const BMM_ENTER_KEY = 2056
Global Const BMM_UP_KEY = 1032
Global Const BMM_DOWN_KEY = 1028
Global Const BMM_LEFT_KEY = 516
Global Const BMM_RIGHT_KEY = 2052
Global Const BMM_F1_KEY = 258
Global Const BMM_F2_KEY = 514
Global Const BMM_F3_KEY = 1026
Global Const BMM_F4_KEY = 2050
Global Const BMM_F5_KEY = 4098
Global Const BMM_F6_KEY = 4104
Global Const BMM_FF1_KEY = 257
Global Const BMM_FF2_KEY = 513
Global Const BMM_FF3_KEY = 1025
Global Const BMM_FF4_KEY = 2049
Global Const BMM_FF5_KEY = 4097
Global Const BMM_NOKEY = 0
Global Const BMM_ONOFF_KEY = 272
Global Const BMM_ABORT_KEY = 776
Global Const BMM_REMOTE_KEY_OFFSET = 8192
Global Const BMM_BEEP_STDINTENS = 100
Global Const BMM_BEEP_STDFREQ = 3900
Global Const BMM_CONTRAST_STDINTENS = 50
Global Const BMM_BACKLIGHT_STDINTENS = 70
Global Const DEL_ERROR = 1
Global Const DEL_TASKSYNCH = 2
Global Const DEL_COMM = 4
Global Const DEL_USER_CMD = 8
Global Const DEL_RESOURCE = 32
Global Const DEL_EXT_INTERFERENCE = 64
Global Const DEL_INTERRUPTION = 128
Global Const DEL_DB_ACCESS = 256
Global Const DEL_FILE_ACCESS = 512
Global Const DEL_MEM_ACCESS = 1024
Global Const DEL_HW_ERROR = 2048
Global Const DEL_ADJ_ERROR = 4096
Global Const DEL_ANG_ERROR = 8192
Global Const DEL_ALL_EVENTS = 16367
Global Const DEL_MAX_LENGTH = 8
Global Const ANG_ERROR = 257
Global Const ANG_INCL_ERROR = 258
Global Const ANG_BAD_ACC = 259
Global Const ANG_BAD_ANGLE_ACC = 260
Global Const ANG_BAD_INCLIN_ACC = 261
Global Const ANG_WRITE_PROTECTED = 266
Global Const ANG_OUT_OF_RANGE = 267
Global Const ANG_IR_OCCURED = 268
Global Const ANG_HZ_MOVED = 269
Global Const ANG_OS_ERROR = 270
Global Const ANG_DATA_ERROR = 271
Global Const ANG_PEAK_CNT_UFL = 272
Global Const ANG_TIME_OUT = 273
Global Const ANG_TOO_MANY_EXPOS = 274
Global Const ANG_PIX_CTRL_ERR = 275
Global Const ANG_MAX_POS_SKIP = 276
Global Const ANG_MAX_NEG_SKIP = 277
Global Const ANG_EXP_LIMIT = 278
Global Const ANG_UNDER_EXPOSURE = 279
Global Const ANG_OVER_EXPOSURE = 280
Global Const ANG_TMANY_PEAKS = 300
Global Const ANG_TLESS_PEAKS = 301
Global Const ANG_PEAK_TOO_SLIM = 302
Global Const ANG_PEAK_TOO_WIDE = 303
Global Const ANG_BAD_PEAKDIFF = 304
Global Const ANG_UNDER_EXP_PICT = 305
Global Const ANG_PEAKS_INHOMOGEN = 306
Global Const ANG_NO_DECOD_POSS = 307
Global Const ANG_UNSTABLE_DECOD = 308
Global Const ANG_TLESS_FPEAKS = 309
Global Const ANG_RAPID_INC = 100
Global Const ANG_STANDARD_INC = 300
Global Const ANG_EXTRA_INC = 900
Global Const ANG_NO_INFO = 0
Global Const ANG_HZ_MOVING = 4096
Global Const ANG_PLANE_ACC = 2
Global Const ANG_ADJ_ACC = 10
Global Const ANG_MAX_LONG = 75
Global Const ANG_MAX_PWD_LEN = 9
Global Const ANG_PEAK_OK = 0
Global Const ANG_WIDTH_OVERFLOW = 1
Global Const ANG_PEAK_UNCOMPL = 2
Global Const ANG_PEAKS_END = 4
Global Const ANG_BAD_WIDTH = 8
Global Const ANG_NO_FINE_PEAK = 16
Global Const ANG_MALUS_SWAP = 32
Global Const ANG_BAD_PDIFF = 64
Global Const RC_COM_ERO = 3072
Global Const RC_COM_CANT_ENCODE = 3073
Global Const RC_COM_CANT_DECODE = 3074
Global Const RC_COM_CANT_SEND = 3075
Global Const RC_COM_CANT_RECV = 3076
Global Const RC_COM_TIMEDOUT = 3077
Global Const RC_COM_WRONG_FORMAT = 3078
Global Const RC_COM_VER_MISMATCH = 3079
Global Const RC_COM_CANT_DECODE_REQ = 3080
Global Const RC_COM_PROC_UNAVAIL = 3081
Global Const RC_COM_CANT_ENCODE_REP = 3082
Global Const RC_COM_SYSTEM_ERR = 3083
Global Const RC_COM_UNKNOWN_HOST = 3084
Global Const RC_COM_FAILED = 3085
Global Const RC_COM_NO_BINARY = 3086
Global Const RC_COM_INTR = 3087
Global Const RC_COM_UNKNOWN_ADDR = 3088
Global Const RC_COM_NO_BROADCAST = 3089
Global Const RC_COM_REQUIRES_8DBITS = 3090
Global Const RC_COM_UD_ERROR = 3091
Global Const RC_COM_LOST_REQ = 3092
Global Const RC_COM_TR_ID_MISMATCH = 3093
Global Const RC_COM_NOT_GEOCOM = 3094
Global Const RC_COM_UNKNOWN_PORT = 3095
Global Const RC_COM_ILLEGAL_TRPT_SELECTOR = 3096
Global Const RC_COM_TRPT_SELECTOR_IN_USE = 3097
Global Const RC_COM_INACTIVE_TRPT_SELECTOR = 3098
Global Const RC_COM_ERO_END = 3099
Global Const RC_COM_OVERRUN = 3100
Global Const RC_COM_SRVR_RX_CHECKSUM_ERROR = 3101
Global Const RC_COM_CLNT_RX_CHECKSUM_ERROR = 3102
Global Const RC_COM_PORT_NOT_AVAILABLE = 3103
Global Const RC_COM_PORT_NOT_OPEN = 3104
Global Const RC_COM_NO_PARTNER = 3105
Global Const RC_COM_ERO_NOT_STARTED = 3106
Global Const RC_COM_CONS_REQ = 3107
Global Const RC_COM_SRVR_IS_SLEEPING = 3108
Global Const RC_COM_SRVR_IS_OFF = 3109
Global Const COM_MAX_ERROR_TXT_LEN = 256
Global Const CSV_NAME_LEN = 20
Global Const CSV_PASSWORD_LEN = 10
Global Const CSV_NUMBER_GEN_PWD = 2
Global Const CSV_USERNAME_LEN = 15
Global Const CSV_SECTION_LEN = 15
Global Const CSV_ENTRY_LEN = 15
Global Const CSV_STRING_LEN = 30
Global Const CSV_DEF_RETICILLUMINTENS = 80
Global Const CSV_DEF_LASERLOTRATE = 50
Global Const CSV_HIGHEST_SERIALNO = 450000
Global Const CSV_LOWEST_SERIALNO = 410000
Global Const RC_CSV_ILLEGAL_USERNR = 4099
Global Const EDM_MAX_MEASMODE = 6
Global Const EDM_COMERR = 769
Global Const EDM_NOSIGNAL = 770
Global Const EDM_PPM_MM = 771
Global Const EDM_METER_FEET = 772
Global Const EDM_ERR12 = 773
Global Const EDM_DIL99 = 774
Global Const EDM_SLDR_TRANSFER_PENDING = 775
Global Const EDM_SLDR_TRANSFER_ILLEGAL = 776
Global Const EDM_SLDR_DATA_ERROR = 777
Global Const EDM_SLDR_CHK_SUM_ERROR = 778
Global Const EDM_SLDR_ADDR_ERROR = 779
Global Const EDM_SLDR_INV_LOADFILE = 780
Global Const EDM_SLDR_UNSUPPORTED = 781
Global Const TMC_NO_FULL_CORRECTION = 1283
Global Const TMC_ACCURACY_GUARANTEE = 1284
Global Const TMC_ANGLE_OK = 1285
Global Const TMC_ANGLE_NO_FULL_CORRECTION = 1288
Global Const TMC_ANGLE_ACCURACY_GUARANTEE = 1289
Global Const TMC_ANGLE_ERROR = 1290
Global Const TMC_DIST_PPM = 1291
Global Const TMC_DIST_ERROR = 1292
Global Const TMC_BUSY = 1293
Global Const TMC_SIGNAL_ERROR = 1294
Global Const TMC_MAX_MEASURE_POINTS = 2
Global Const TMC_PT_LAST = 1
Global Const TMC_MEASURE = 0
Global Const TMC_LAST_MEASURING = 1
Global Const TMC_1_BEFORE_LAST_MEASURING = 2
Global Const TMC_ANGLE_VALID_TIME = 700
Global Const MOT_AXES = 2
Global Const MOT_RC_UNREADY = 1792
Global Const MOT_RC_BUSY = 1793
Global Const MOT_RC_NOT_OCONST = 1794
Global Const MOT_RC_NOT_CONFIG = 1795
Global Const MOT_RC_NOT_POSIT = 1796
Global Const MOT_RC_NOT_SERVICE = 1797
Global Const MOT_RC_NOT_BUSY = 1798
Global Const MOT_RC_NOT_LOCK = 1799
Global Const MAX_FIL_BUFFER = 450
Global Const FIL_LEN_NAME = 8
Global Const FIL_LEN_EXTN = 3
Global Const FIL_LEN_LABEL = 11
Global Const FIL_LEN_PATH = 64
Global Const FIL_MAX_OPENFILES = 20
Global Const FIL_LEN_RSRVD = 1
Global Const FIL_ACC_RDONLY = 0
Global Const FIL_ACC_WRONLY = 1
Global Const FIL_ACC_RDWR = 2
Global Const FIL_ACC_APPEND = 8
Global Const FIL_ACC_CREATE = 256
Global Const FIL_ACC_OVWRIT = 512
Global Const FIL_ACC_EXCL = 1024
Global Const FIL_ATT_NONE = 0
Global Const FIL_ATT_NORMAL = 0
Global Const FIL_ATT_RDONLY = 1
Global Const FIL_ATT_HIDDEN = 2
Global Const FIL_ATT_SYSTEM = 4
Global Const FIL_ATT_VOLNAM = 8
Global Const FIL_ATT_SUBDIR = 16
Global Const FIL_ATT_ARCHIV = 32
Global Const RC_FIL_NO_ERROR = 3840
Global Const RC_FIL_FILNAME_NOT_FOUND = 3845
Global Const RC_FIL_NO_MAKE_DIRECTORY = 3880
Global Const RC_FIL_RENAME_FILE_FAILED = 3886
Global Const RC_FIL_INVALID_PATH = 3888
Global Const RC_FIL_FILE_NOT_DELETED = 3898
Global Const RC_FIL_ILLEGAL_ORIGIN = 3906
Global Const RC_FIL_END_OF_FILE = 3924
Global Const RC_FIL_NO_MORE_ROOM_ON_MEDIUM = 3931
Global Const RC_FIL_PATTERN_DOES_NOT_MATCH = 3932
Global Const RC_FIL_FILE_ALREADY_OPEND_FOR_01 = 3948
Global Const RC_FIL_WRITE_TO_MEDIUM_FAILED = 3957
Global Const RC_FIL_START_SEARCH_NOT_CALLED = 3963
Global Const RC_FIL_NO_STORAGE_MEDIUM_IN_DE02 = 3964
Global Const RC_FIL_ILLEGAL_FILE_OPEN_TYPE = 3965
Global Const RC_FIL_MEDIUM_NEWLY_INSERTED = 3966
Global Const RC_FIL_MEMORY_FAILED = 3967
Global Const RC_FIL_FATAL_ERROR = 3968
Global Const RC_FIL_FAT_ERROR = 3969
Global Const RC_FIL_ILLEGAL_DRIVE = 3970
Global Const RC_FIL_INVALID_FILE_DESCR = 3971
Global Const RC_FIL_SEEK_FAILED = 3972
Global Const RC_FIL_CANNOT_DELETE = 3973
Global Const RC_FIL_MEDIUM_WRITE_PROTECTED = 3974
Global Const RC_FIL_BATTERY_LOW = 3975
Global Const RC_FIL_BAD_FORMAT = 3976
Global Const AUT_RC_TIMEOUT = 8704
Global Const AUT_RC_DETENT_ERROR = 8705
Global Const AUT_RC_ANGLE_ERROR = 8706
Global Const AUT_RC_MOTOR_ERROR = 8707
Global Const AUT_RC_INCACC = 8708
Global Const AUT_RC_DEV_ERROR = 8709
Global Const AUT_RC_NO_TARGET = 8710
Global Const AUT_RC_MULTIPLE_TARGETS = 8711
Global Const AUT_RC_BAD_ENVIRONMENT = 8712
Global Const AUT_RC_DETECTOR_ERROR = 8713
Global Const AUT_RC_NOT_ENABLED = 8714
Global Const AUT_RC_CALACC = 8715
Global Const AUT_RC_ACCURACY = 8716
Global Const AUT_SEARCH_DEF = 0.04
Global Const TXT_NAME_LENGTH = 20
Global Const TXT_REF_LANGUAGE = 0
Global Const TXT_MAX_LANGUAGE = 2
Global Const TXT_MAX_TOKEN = 20000
Global Const TXT_SYST_8 = 8
Global Const TXT_NIL_TOKEN = -1
Global Const TXT_GROUPSIZE = 1000
Global Const TXT_OTHER_LANG = 2560
Global Const TXT_UNDEF_TOKEN = 2561
Global Const TXT_UNDEF_LANG = 2562
Global Const TXT_TOOMANY_LANG = 2563
Global Const TXT_GROUP_OCC = 2564
Global Const TXT_INVALID_GROUP = 2565
Global Const TXT_OUT_OF_MEM = 2566
Global Const TXT_MEM_ERROR = 2567
Global Const TXT_TRANSFER_PENDING = 2568
Global Const TXT_TRANSFER_ILLEGAL = 2569
Global Const TXT_INVALID_SIZE = 2570
Global Const TXT_ALREADY_EXIST = 2571
Global Const MEM_OUT_OF_MEMORY = 1536
Global Const MEM_OUT_OF_HANDLES = 1537
Global Const MEM_TAB_OVERFLOW = 1538
Global Const MEM_HANDLE_INVALID = 1539
Global Const MEM_DATA_NOT_FOUND = 1540
Global Const MEM_DELETE_ERROR = 1541
Global Const MEM_ZERO_ALLOC_ERR = 1542
Global Const MEM_REORG_ERR = 1543
Global Const LDR_NOT_USED = -1
Global Const LDR_DATABUFFER_SIZE = 450
Global Const LDR_MAX_LEVELS = 10
Global Const LDR_LEVEL_LEN = 30
Global Const LDR_MAX_NAME_LEN = 20
Global Const LDR_MAX_ANZ_FKT = 20
Global Const LDR_MAX_SECTIONS = 15
Global Const LDR_MAX_SEC_NAME_LEN = 10
Global Const LDR_PENDING = 2048
Global Const LDR_PRGM_OCC = 2049
Global Const LDR_TRANSFER_ILLEGAL = 2050
Global Const LDR_NOT_FOUND = 2051
Global Const LDR_ALREADY_EXIST = 2052
Global Const LDR_NOT_EXIST = 2053
Global Const LDR_SIZE_ERROR = 2054
Global Const LDR_MEM_ERROR = 2055
Global Const LDR_PRGM_NOT_EXIST = 2056
Global Const LDR_FUNC_LEVEL_ERR = 2057
Global Const LDR_RECURSIV_ERR = 2058
Global Const LDR_INST_ERR = 2059
Global Const LDR_FUNC_OCC = 2060
Global Const LDR_RUN_ERROR = 2061
Global Const LDR_DEL_MENU_ERR = 2062
Global Const LDR_OBJ_TYPE_ERROR = 2063
Global Const LDR_WRONG_SECKEY = 2064
Global Const LDR_ILLEGAL_LOADADR = 2065
Global Const LDR_IEEE_ERROR = 2066
Global Const LDR_WRONG_APPL_VERSION = 2067
Global Const BMM_CHAR_WIDTH = 6
Global Const BMM_CHAR_HEIGHT = 8
Global Const BMM_NO_OF_COLS = 35
Global Const BMM_NO_OF_ROWS = 8
Global Const BMM_NO_OF_GRAPHIC_COLS = 210
Global Const BMM_NO_OF_GRAPHIC_ROWS = 64
Global Const FONT_NAME_LEN = 20
Global Const CHARSET_GEO = 0
Global Const CHARSET_SYMB = 1
Global Const CHARSET_ARAB_CLASS = 200
Global Const CHARSET_ARAB = 200
Global Const BMM_XFER_PENDING = 2305
Global Const BMM_NO_XFER_OPEN = 2306
Global Const BMM_UNKNOWN_CHARSET = 2307
Global Const BMM_NOT_INSTALLED = 2308
Global Const BMM_ALREADY_EXIST = 2309
Global Const BMM_CANT_DELETE = 2310
Global Const BMM_MEM_ERROR = 2311
Global Const BMM_CHARSET_USED = 2312
Global Const BMM_CHARSET_SAVED = 2313
Global Const BMM_INVALID_ADR = 2314
Global Const BMM_CANCELANDADR_ERROR = 2315
Global Const BMM_INVALID_SIZE = 2316
Global Const BMM_CANCELANDINVSIZE_ERROR = 2317
Global Const BMM_ALL_GROUP_OCC = 2318
Global Const BMM_CANT_DEL_LAYERS = 2319
Global Const BMM_UNKNOWN_LAYER = 2320
Global Const BMM_INVALID_LAYERLEN = 2321
Global Const BMM_REFRESH_KEY = 16383
Global Const ATA_MAX_PWD_LEN = 9
Global Const ATA_MAX_DEVTYPE_LEN = 9
Global Const ATA_RC_NOT_READY = 512
Global Const ATA_RC_NO_RESULT = 513
Global Const ATA_RC_SEVERAL_TARGETS = 514
Global Const ATA_RC_BIG_SPOT = 515
Global Const ATA_RC_BACKGROUND = 516
Global Const ATA_RC_NO_TARGETS = 517
Global Const ATA_RC_NOT_ACCURAT = 518
Global Const ATA_RC_SPOT_ON_EDGE = 519
Global Const ATA_RC_SETTINGS_INCOMPLETE = 520
Global Const ATA_RC_BLOOMING = 522
Global Const ATA_RC_NOT_BUSY = 523
Global Const ATA_RC_STRANGE_LIGHT = 524
Global Const ATA_RC_V24_FAIL = 525
Global Const ATA_RC_DECODE_ERROR = 526
Global Const ATA_RC_HZ_FAIL = 527
Global Const ATA_RC_V_FAIL = 528
Global Const ATA_RC_HZ_STRANGE_L = 529
Global Const ATA_RC_V_STRANGE_L = 530
Global Const ATA_SLDR_TRANSFER_PENDING = 531
Global Const ATA_SLDR_TRANSFER_ILLEGAL = 532
Global Const ATA_SLDR_DATA_ERROR = 533
Global Const ATA_SLDR_CHK_SUM_ERROR = 534
Global Const ATA_SLDR_ADDRESS_ERROR = 535
Global Const ATA_SLDR_INV_LOADFILE = 536
Global Const ATA_SLDR_UNSUPPORTED = 537
Global Const NO_SY1_TOKEN = 50
Global Const BASE_SY1_TOKEN = 1000
Global Const K_ABORT_TOK = 1000
Global Const K_END_TOK = 1001
Global Const K_HELP_TOK = 1002
Global Const K_NO_TOK = 1003
Global Const K_OK_TOK = 1004
Global Const K_YES_TOK = 1005
Global Const U_PERCENT_TOK = 1006
Global Const U_PERCENT0_TOK = 1007
Global Const U_PERCENT100_TOK = 1008
Global Const K_DEFAULT_TOK = 1009
Global Const K_EDIT_TOK = 1010
Global Const K_OFF_TOK = 1011
Global Const K_INSERT_TOK = 1012
Global Const K_LL_ARR_TOK = 1013
Global Const K_RR_ARR_TOK = 1014
Global Const K_L_ARR_TOK = 1015
Global Const K_R_ARR_TOK = 1016
Global Const NO_HELP_TOK = 1017
Global Const K_TXT_EDIT_TOK = 1018
Global Const K_NUM_EDIT_TOK = 1019
Global Const K_LIST_TOK = 1020
Global Const K_BOOL_ON_TOK = 1021
Global Const K_BOOL_OFF_TOK = 1022
Global Const K_BOOL_YES_TOK = 1023
Global Const K_BOOL_NO_TOK = 1024
Global Const K_BOOL_CHK_TOK = 1025
Global Const K_BOOL_NOCHK_TOK = 1026
Global Const K_SPIN_TOK = 1027
Global Const K_CONT_TOK = 1028
Global Const K_RETRY_TOK = 1029
Global Const T_HELP_CAPT_TOK = 1030
Global Const K_BOOL_WITH_TOK = 1031
Global Const K_BOOL_WITHOUT_TOK = 1032
Global Const K_BOOL_OLD_TOK = 1033
Global Const K_BOOL_NEW_TOK = 1034
Global Const U_MM_TOK = 1035
Global Const U_NM_TOK = 1036
Global Const U_KBYTE_TOK = 1037
Global Const K_R_END_ARR_TOK = 1038
Global Const K_L_END_ARR_TOK = 1039
Global Const K_IGNORE_TOK = 1040
Global Const U_BYTE_TOK = 1041
Global Const K_BOOL_NEW_SET_TOK = 1042
Global Const K_BOOL_NO_SET_TOK = 1043
Global Const K_BOOL_OFF_SET_TOK = 1044
Global Const K_BOOL_OLD_SET_TOK = 1045
Global Const K_BOOL_ON_SET_TOK = 1046
Global Const K_BOOL_WITH_SET_TOK = 1047
Global Const K_BOOL_WITHOUT_SET_TOK = 1048
Global Const K_BOOL_YES_SET_TOK = 1049
Global Const MMI_MSG_TO_ENDLESS = -1
Global Const MMI_MSG_TO_STANDARD = 3000
Global Const MMI_MSG_TO_LONG = 30000
Global Const MMI_MAX_MSGBOX_BUTTONS = 6
Global Const MMI_MAX_FIXLINES = 3
Global Const FIELD_ALIGN_MASK = 7
Global Const MMI_FIELD_RIGHT = 0
Global Const MMI_FIELD_LEFT = 1
Global Const MMI_FIELD_CENTER = 2
Global Const FIELD_DELI_MASK = 48
Global Const MMI_FIELD_COLON = 16
Global Const MMI_FIELD_VERTBAR = 32
Global Const MMI_FIELD_INOUT_MASK = 7
Global Const MMI_FIELD_INVISIBLE = 0
Global Const MMI_FIELD_OUT = 1
Global Const MMI_FIELD_INOUT = 2
Global Const MMI_FIELD_UFUNC_MASK = 8
Global Const MMI_FIELD_UFUNC_ENTER = 0
Global Const MMI_FIELD_UFUNC_KEY = 8
Global Const MMI_FIELD_DEFINE_MASK = 112
Global Const MMI_VALUE_NOT_DEFINED = 0
Global Const MMI_VALUE_TOO_LARGE = 16
Global Const MMI_VALUE_VALID = 32
Global Const MMI_VALUE_CLIPPED = 128
Global Const MMI_LEVEL_NOT_DEF = 0
Global Const MMI_LEVEL_REDUCED_USER = 4
Global Const MMI_LEVEL_FULL_USER = 8
Global Const MAX_MENU_ITEMS = 50
Global Const MMI_MENU_LAST = 50
Global Const MMI_BUTTON_ID_EXISTS = 2817
Global Const MMI_DLG_NOT_OPEN = 2818
Global Const MMI_DLG_OPEN = 2819
Global Const MMI_DLG_SPEC_MISMATCH = 2820
Global Const MMI_DLGDEF_EMPTY = 2821
Global Const MMI_DLGDEF_NOT_OPEN = 2822
Global Const MMI_DLGDEF_OPEN = 2823
Global Const MMI_FIELD_ID_EXISTS = 2824
Global Const MMI_ILLEGAL_APP_ID = 2825
Global Const MMI_ILLEGAL_BUTTON_ID = 2826
Global Const MMI_ILLEGAL_DLG_ID = 2827
Global Const MMI_ILLEGAL_FIELD_COORDS = 2828
Global Const MMI_ILLEGAL_FIELD_ID = 2829
Global Const MMI_ILLEGAL_FIELD_TYPE = 2830
Global Const MMI_ILLEGAL_FIELD_FORMAT = 2831
Global Const MMI_ILLEGAL_FIXLINES = 2832
Global Const MMI_ILLEGAL_MB_TYPE = 2833
Global Const MMI_ILLEGAL_MENU_ID = 2834
Global Const MMI_ILLEGAL_MENUITEM_ID = 2835
Global Const MMI_ILLEGAL_NEXT_ID = 2836
Global Const MMI_ILLEGAL_TOPLINE = 2837
Global Const MMI_NOMORE_BUTTONS = 2838
Global Const MMI_NOMORE_DLGS = 2839
Global Const MMI_NOMORE_FIELDS = 2840
Global Const MMI_NOMORE_MENUS = 2841
Global Const MMI_NOMORE_MENUITEMS = 2842
Global Const MMI_NOMORE_WINDOWS = 2843
Global Const MMI_SYS_BUTTON = 2844
Global Const MMI_VREF_UNDEF = 2845
Global Const MMI_EXIT_DLG = 2846
Global Const MMI_KEEP_FOCUS = 2847
Global Const MMI_NOMORE_ITEMS = 2848
Global Const MMI_NO_FOCUS = 255
Global Const MMI_NO_FIELD = 255
Global Const RC_DPL_IVRESULT = 4865
Global Const RC_DPL_CHANGED = 4866
Global Const RC_DPL_SAVED = 4867
Global Const RC_DPL_PROJ_ERR = 4868
Global Const RC_DPL_NOT_FOUND = 4869
Global Const RC_DPL_IV_RELEASE = 4870
Global Const RC_DPL_IV_PROJECT = 4871
Global Const RC_DPL_IV_JOB_INFO = 4872
Global Const RC_DPL_IV_PNTINFO = 4873
Global Const RC_DPL_IV_PNTATTRIB = 4874
Global Const RC_DPL_IV_COORD = 4875
Global Const RC_DPL_IV_MEASURE = 4876
Global Const RC_DPL_IV_INSTR_DATA = 4877
Global Const RC_DPL_IV_INSTR_CONFIG = 4878
Global Const RC_DPL_IV_METEO = 4879
Global Const RC_DPL_IV_LINE = 4880
Global Const RC_DPL_IV_TMCINFO = 4881
Global Const RC_DPL_REC_NOT_ALLOWED = 4882
Global Const DPL_STATION_POINT = 0
Global Const DPL_CURRENT_POINT = 1
Global Const DPL_LAST_POINT = 2
Global Const DPL_POINT_1 = 3
Global Const DPL_MAIN_POINT = 4
Global Const DPL_MAX_DATA_SET = 5
Global Const DPL_STATION_CODE = 0
Global Const DPL_TARGET_CODE = 1
Global Const DPL_LINE_CODE = 2
Global Const DPL_MAX_CODE_SET = 3
Global Const DIALOGUE_LEN = 16
Global Const WIR_PTNR_OVERFLOW = 5121
Global Const WIR_NUM_ASCII_CARRY = 5122
Global Const WIR_PTNR_NO_INC = 5123
Global Const WIR_STEP_SIZE = 5124
Global Const WIR_BUSY = 5125
Global Const WIR_CONFIG_FNC = 5127
Global Const WIR_CANT_OPEN_FILE = 5128
Global Const WIR_FILE_WRITE_ERROR = 5129
Global Const WIR_MEDIUM_NOMEM = 5130
Global Const WIR_NO_MEDIUM = 5131
Global Const WIR_EMPTY_FILE = 5132
Global Const WIR_INVALID_DATA = 5133
Global Const WIR_F2_BUTTON = 5134
Global Const WIR_F3_BUTTON = 5135
Global Const WIR_F4_BUTTON = 5136
Global Const WIR_SHF2_BUTTON = 5137
Global Const WIR_ASCII = 1
Global Const WIR_ASCII_SIGN = 2
Global Const WIR_DOUBLE = 3
Global Const WIR_GSI_LEN = 16
Global Const WIR_ARRAY_LEN = 55
Global Const WIR_DLG_FIX_LINE = 1
Global Const WIR_ASCII_STRLEN_GSI8 = 10
Global Const WIR_ASCII_STRLEN = 18
Global Const WIR_MAX_REC_WI = 12
Global Const WIR_DLG_LINES = 11
Global Const GSI_WORK_PATH_LENGTH = 128
Global Const GSI_WORK_FILENAME_LENGTH = 13
Global Const WIR_MAX_NO_FILES = 12
Global Const WIR_MAX_NO_ALL_FILES = 24
Global Const WIR_REG_AXES = 2
Global Const WIR_NO_EXTRA_CHARS = 5
Global Const WIR_RECFORMAT_GSI = 0
Global Const WIR_RECFORMAT_GSI16 = 1
Global Const WIR_DISPFORMAT_GSIMEAS = 11
Global Const WIR_DISPFORMAT_GSIALL = 12
Global Const WIR_NONE_KEY = 0
Global Const WIR_F1_KEY = 1
Global Const WIR_F2_KEY = 2
Global Const WIR_F3_KEY = 4
Global Const WIR_F4_KEY = 8
Global Const WIR_F5_KEY = 22
Global Const WIR_F6_KEY = 50
Global Const BAP_CHANGE_ALL_TO_DIST = 9217
Global Const BAP_SD_TEXT1 = 0
Global Const BAP_SD_TEXT2 = 1
Global Const BAP_SD_DESCR1 = 2
Global Const BAP_SD_VAL1 = 3
Global Const BAP_SD_DIM1 = 4
Global Const BAP_SD_DESCR2 = 5
Global Const BAP_SD_VAL2 = 6
Global Const BAP_SD_DIM2 = 7
Global Const BAP_SD_NO_FIELDS2 = 8
Global Const BAP_SD_NO_FIELDS1 = 5
Global Const BAP_SD_FIXLINES = 0
Global Const BAP_MD_VAL1 = 1
Global Const BAP_MD_VAL2 = 4
Global Const BAP_MD_VAL3 = 7
Global Const BAP_MD_VAL4 = 10
Global Const BAP_MD_VAL5 = 13
Global Const BAP_MD_VAL6 = 16
Global Const BAP_MD_VAL7 = 19
Global Const BAP_MD_VAL8 = 22
Global Const BAP_MD_VAL9 = 27
Global Const BAP_MD_VAL10 = 30
Global Const BAP_MD_VAL11 = 33
Global Const BAP_MAX_MEASDLG_LINES = 12
Global Const BAP_NO_MSG_NR = -1
Global Const BAP_DEF_MSG_NR = 0
Global Const BAPL_NUM_OF_PIXEL = 1440
Global Const BAPL_MAX_LIST_ELEMENTS = 25
Global Const BAPL_LEN_LIST_ELEMENT = 30
Global Const BAP_DEFAULT = 0
Global Const BAP_DIM_ON = 1
Global Const BAP_SPECIALKEYS_ON = 2
Global Const BAP_ENTRIES_TOT = 71
Global Const BAP_ENTRY_STATION = 0
Global Const BAP_ENTRY_TARGET = 12
Global Const BAP_ENTRY_MEAS = 28
Global Const BAP_ENTRY_VARIA = 42
Global Const BAP_ENTRY_OTHER = 69
Global Const BAP_ENTRY_GROUPMAX = 26
Global Const BEEP_NO = 0
Global Const BEEP_LOW = 20
Global Const BEEP_HIGH = 100
Global Const AUS_LASERLOT_TEMP_ERR = 8960

' Enumerated global constants

' Enumerated type BOOLE

Global Const FALSE_ = 0
Global Const TRUE_ = 1

' Enumerated type ON_OFF_TYPE

Global Const OFF = 0
Global Const ON_ = 1

' Enumerated type TPS_DEVICE_TYPE

Global Const TPS_DEVICE_T = 0
Global Const TPS_DEVICE_TC1 = 1
Global Const TPS_DEVICE_TC2 = 2
Global Const TPS_DEVICE_MOT = 4
Global Const TPS_DEVICE_ATR = 8
Global Const TPS_DEVICE_EGL = 16
Global Const TPS_DEVICE_DB = 32
Global Const TPS_DEVICE_DL = 64
Global Const TPS_DEVICE_LP = 128
Global Const TPS_DEVICE_SIM = 16384

' Enumerated type TPS_DEVICE_CLASS

Global Const TPS_CLASS_1100 = 0
Global Const TPS_CLASS_1700 = 1
Global Const TPS_CLASS_1800 = 2
Global Const TPS_CLASS_5000 = 3
Global Const TPS_CLASS_6000 = 4
Global Const TPS_CLASS_1500 = 5

' Enumerated type COM_TPS_STARTUP_MODE

Global Const COM_TPS_REMOTE = 1

' Enumerated type COM_TPS_STOP_MODE

Global Const COM_TPS_SHUT_DOWN = 0
Global Const COM_TPS_SLEEP = 1

' Enumerated type EDM_MODE

Global Const EDM_SINGLE_STANDARD = 0
Global Const EDM_SINGLE_EXACT = 1
Global Const EDM_SINGLE_FAST = 2
Global Const EDM_CONT_STANDARD = 3
Global Const EDM_CONT_EXACT = 4
Global Const EDM_CONT_FAST = 5
Global Const EDM_UNDEFINED = 6

' Enumerated type EDM_TRKLIGHT_BRIGHTNESS

Global Const EDM_LOW_BRIGHTNESS = 0
Global Const EDM_MEDIUM_BRIGHTNESS = 1
Global Const EDM_HIGH_BRIGHTNESS = 2

' Enumerated type TMC_FACE

Global Const TMC_FACE_1 = 0
Global Const TMC_FACE_2 = 1

' Enumerated type TMC_INCLINE_PRG

Global Const TMC_MEA_INC = 0
Global Const TMC_AUTO_INC = 1
Global Const TMC_PLANE_INC = 2

' Enumerated type TMC_MEASURE_PRG

Global Const TMC_STOP = 0
Global Const TMC_DEF_DIST = 1
Global Const TMC_TRK_DIST = 2
Global Const TMC_CLEAR = 3
Global Const TMC_SIGNAL = 4
Global Const TMC_RTRK_DIST = 8

' Enumerated type MOT_MODE

Global Const MOT_POSIT = 0
Global Const MOT_OCONST = 1
Global Const MOT_MANUPOS = 2
Global Const MOT_LOCK = 3
Global Const MOT_BREAK = 4
Global Const MOT_TERM = 7

' Enumerated type MOT_LOCK_STATUS

Global Const MOT_LOCKED_OUT = 0
Global Const MOT_LOCKED_IN = 1
Global Const MOT_PREDICTION = 2

' Enumerated type MOT_STOPMODE

Global Const MOT_NORMAL = 0
Global Const MOT_SHUTDOWN = 1

' Enumerated type AUT_POSMODE

Global Const AUT_NORMAL = 0
Global Const AUT_PRECISE = 1

' Enumerated type AUT_ATRMODE

Global Const AUT_POSITION = 0
Global Const AUT_TARGET = 1

' Enumerated type AUT_ADJMODE

Global Const AUT_NORM_MODE = 0
Global Const AUT_POINT_MODE = 1

' Enumerated type SUP_AUTO_POWER

Global Const AUTO_POWER_DISABLED = 0
Global Const AUTO_POWER_SLEEP = 1
Global Const AUTO_POWER_OFF = 2

' Enumerated type BAP_MEASURE_PRG

Global Const BAP_NO_MEAS = 0
Global Const BAP_NO_DIST = 1
Global Const BAP_DEF_DIST = 2
Global Const BAP_TRK_DIST = 3
Global Const BAP_RTRK_DIST = 4
Global Const BAP_CLEAR_DIST = 5
Global Const BAP_STOP_TRK = 6

' Structure definitions

Type TPS_DEVICE
    eClass As Long                                ' TPS_DEVICE_CLASS
    Type As Long                                  ' TPS_DEVICE_TYPE
End Type

Type DATE_TYPE
    Year As Integer
    Month As Byte
    Day As Byte
End Type

Type TIME_TYPE
    Hour As Byte
    Minute As Byte
    Second As Byte
    cPadding As Byte
End Type

Type DATIME
    date As DATE_TYPE
    Time As TIME_TYPE
End Type

Type TMC_INCLINE
    dCrossIncline As Double
    dLengthIncline As Double
    dAccuracyIncline As Double
    InclineTime As Long
End Type

Type TMC_ANGLE
    dHz As Double
    dV As Double
    dAngleAccuracy As Double
    AngleTime As Long
    Incline As TMC_INCLINE
    eFace As Long                                 ' TMC_FACE
End Type

Type TMC_HZ_V_ANG
    dHz As Double
    dV As Double
End Type

Type TMC_COORDINATE
    dE As Double
    dN As Double
    dH As Double
    CoordTime As Long
    dE_Cont As Double
    dN_Cont As Double
    dH_Cont As Double
    CoordContTime As Long
End Type

Type TMC_EDM_SIGNAL
    dSignalIntensity As Double
    Time As Long
End Type

Type TMC_ANG_SWITCH
    eInclineCorr As Long                          ' ON_OFF_TYPE
    eStandAxisCorr As Long                        ' ON_OFF_TYPE
    eCollimationCorr As Long                      ' ON_OFF_TYPE
    eTiltAxisCorr As Long                         ' ON_OFF_TYPE
End Type

Type TMC_ATMOS_TEMPERATURE
    dLambda As Double
    dPressure As Double
    dDryTemperature As Double
    dWetTemperature As Double
End Type

Type TMC_REFRACTION
    eRefOn As Long                                ' ON_OFF_TYPE
    dEarthRadius As Double
    dRefractiveScale As Double
End Type

Type TMC_HEIGHT
    dHr As Double
End Type

Type TMC_STATION
    dE0 As Double
    dN0 As Double
    dH0 As Double
    dHi As Double
End Type

Type MOT_CONTROLPAR
    adPAngle(0 To 1) As Double
    adIAngle(0 To 1) As Double
    adPOmega(0 To 1) As Double
    adIOmega(0 To 1) As Double
    dAngleTime As Double
    dOmegaTime As Double
End Type

Type MOT_MOTIONPAR
    adSpeedLimit(0 To 1) As Double
    adAccel(0 To 1) As Double
End Type

Type MOT_MAXPAR
    adMaxSpeed(0 To 1) As Double
    adMaxAccel(0 To 1) As Double
End Type

Type MOT_POSTOL
    adPosTol(0 To 1) As Double
End Type

Type MOT_KNOBPAR
    alGainLevel3(0 To 1) As Long
    anGainLevel2(0 To 1) As Integer
    anGainLevel1(0 To 1) As Integer
    anStep2(0 To 1) As Integer
    anStep1(0 To 1) As Integer
    anKickDown(0 To 1) As Integer
End Type

Type MOT_COM_PAIR
    adValue(0 To 1) As Double
End Type

Type AUT_POSTOL
    adPosTol(0 To 1) As Double
End Type

Type AUT_TIMEOUT
    adPosTimeout(0 To 1) As Double
End Type

Type BMM_ICON_DESC
    unCol As Integer
    unRow As Integer
    unWidth As Integer
    unHeight As Integer
    ucSymbol As Byte
End Type

' Typedef'd arrays

' Geocom Visual Basic API

' Enumerated type COM_PORT

Global Const COM_1 = 0
Global Const COM_2 = 1
Global Const COM_3 = 2
Global Const COM_4 = 3

' Enumerated type COM_FORMAT

Global Const COM_ASCII = 0
Global Const COM_BINARY = 1

' Enumerated type COM_BAUD_RATE

Global Const COM_BAUD_38400 = 0
Global Const COM_BAUD_19200 = 1
Global Const COM_BAUD_9600 = 2
Global Const COM_BAUD_4800 = 3
Global Const COM_BAUD_2400 = 4

' Enumerated type COM_TPS_STATUS

Global Const COM_TPS_OFF = 0
Global Const COM_TPS_SLEEPING = 1
Global Const COM_TPS_READY = 2
Global Const COM_TPS_UNKNOWN = 3

' Geocom public local functions

Declare Function VB_COM_Init Lib "GEOCOM32.DLL" () As Integer
Declare Function VB_COM_End Lib "GEOCOM32.DLL" () As Integer
Declare Function VB_COM_OpenConnection Lib "GEOCOM32.DLL" (ByVal Port As Long, ByVal Baud As Long, ByVal Retries As Integer) As Integer
Declare Function VB_COM_CloseConnection Lib "GEOCOM32.DLL" () As Integer
Declare Function VB_COM_GetBaudRate Lib "GEOCOM32.DLL" (eRate As Long) As Integer
Declare Function VB_COM_SetBaudRate Lib "GEOCOM32.DLL" (ByVal eRate As Long) As Integer
Declare Function VB_COM_GetDoublePrecision Lib "GEOCOM32.DLL" (nDigits As Integer) As Integer
Declare Function VB_COM_SetDoublePrecision Lib "GEOCOM32.DLL" (ByVal nDigits As Integer) As Integer
Declare Function VB_COM_GetComFormat Lib "GEOCOM32.DLL" (Format As Long) As Integer
Declare Function VB_COM_SetComFormat Lib "GEOCOM32.DLL" (ByVal Format As Long) As Integer
Declare Function VB_COM_GetTimeOut Lib "GEOCOM32.DLL" (Timeout As Integer) As Integer
Declare Function VB_COM_SetTimeOut Lib "GEOCOM32.DLL" (ByVal Timeout As Integer) As Integer
Declare Function VB_COM_GetTPSState Lib "GEOCOM32.DLL" (State As Long) As Integer
Declare Function VB_COM_UseWindow Lib "GEOCOM32.DLL" (ByVal hWnd As Integer) As Integer
Declare Function VB_COM_SetConnDlgFlag Lib "GEOCOM32.DLL" (ByVal bShow As Long) As Integer
Declare Function VB_COM_SetERODlgFlag Lib "GEOCOM32.DLL" (ByVal bShow As Long) As Integer
Declare Function VB_COM_RunERO Lib "GEOCOM32.DLL" (RetCode As Integer) As Integer
Declare Function VB_COM_Abort Lib "GEOCOM32.DLL" () As Integer
Declare Function VB_COM_ViewError Lib "GEOCOM32.DLL" (ByVal RetCode As Integer, ByVal szMsgTitle As String) As Integer
Declare Function VB_COM_GetErrorText Lib "GEOCOM32.DLL" (ByVal RetCode As Integer, ByVal szErrorText As String) As Integer
Declare Function VB_COM_GetWinSWVersion Lib "GEOCOM32.DLL" (nRel As Integer, nVer As Integer, nSubVer As Integer) As Integer

' Remote functions

Declare Function VB_COM_NullProc Lib "GEOCOM32.DLL" () As Integer
Declare Function VB_COM_Local Lib "GEOCOM32.DLL" () As Integer
Declare Function VB_COM_SetSendDelay Lib "GEOCOM32.DLL" (in1 As Integer) As Integer
Declare Function VB_COM_GetSWVersion Lib "GEOCOM32.DLL" (out1 As Integer, out2 As Integer, out3 As Integer) As Integer
Declare Function VB_COM_SwitchOnTPS Lib "GEOCOM32.DLL" (in1 As Long) As Integer
Declare Function VB_COM_SwitchOffTPS Lib "GEOCOM32.DLL" (in1 As Long) As Integer
Declare Function VB_COM_GetBinaryAvailable Lib "GEOCOM32.DLL" (out1 As Long) As Integer
Declare Function VB_COM_SetBinaryAvailable Lib "GEOCOM32.DLL" (in1 As Long) As Integer
Declare Function VB_COM_EnableSignOff Lib "GEOCOM32.DLL" (in1 As Long) As Integer
Declare Function VB_EDM_Laserpointer Lib "GEOCOM32.DLL" (in1 As Long) As Integer
Declare Function VB_EDM_SetBumerang Lib "GEOCOM32.DLL" (in1 As Long) As Integer
Declare Function VB_EDM_On Lib "GEOCOM32.DLL" (in1 As Long) As Integer
Declare Function VB_EDM_SetTrkLightSwitch Lib "GEOCOM32.DLL" (in1 As Long) As Integer
Declare Function VB_EDM_SetTrkLightBrightness Lib "GEOCOM32.DLL" (in1 As Long) As Integer
Declare Function VB_EDM_GetTrkLightSwitch Lib "GEOCOM32.DLL" (out1 As Long) As Integer
Declare Function VB_EDM_GetTrkLightBrightness Lib "GEOCOM32.DLL" (out1 As Long) As Integer
Declare Function VB_EDM_GetBumerang Lib "GEOCOM32.DLL" (out1 As Long) As Integer
Declare Function VB_TMC_GetAngle1 Lib "GEOCOM32.DLL" (out1 As TMC_ANGLE, in2 As Long) As Integer
Declare Function VB_TMC_SetInclineSwitch Lib "GEOCOM32.DLL" (in1 As Long) As Integer
Declare Function VB_TMC_GetInclineSwitch Lib "GEOCOM32.DLL" (out1 As Long) As Integer
Declare Function VB_TMC_DoMeasure Lib "GEOCOM32.DLL" (in1 As Long, in2 As Long) As Integer
Declare Function VB_TMC_GetStation Lib "GEOCOM32.DLL" (out1 As TMC_STATION) As Integer
Declare Function VB_TMC_SetStation Lib "GEOCOM32.DLL" (in1 As TMC_STATION) As Integer
Declare Function VB_TMC_GetHeight Lib "GEOCOM32.DLL" (out1 As TMC_HEIGHT) As Integer
Declare Function VB_TMC_SetHeight Lib "GEOCOM32.DLL" (in1 As TMC_HEIGHT) As Integer
Declare Function VB_TMC_GetAngSwitch Lib "GEOCOM32.DLL" (out1 As TMC_ANG_SWITCH) As Integer
Declare Function VB_TMC_SetAngSwitch Lib "GEOCOM32.DLL" (in1 As TMC_ANG_SWITCH) As Integer
Declare Function VB_TMC_SetHandDist Lib "GEOCOM32.DLL" (in1 As Double, in2 As Double, in3 As Long) As Integer
Declare Function VB_TMC_SetEdmMode Lib "GEOCOM32.DLL" (in1 As Long) As Integer
Declare Function VB_TMC_GetEdmMode Lib "GEOCOM32.DLL" (out1 As Long) As Integer
Declare Function VB_TMC_GetSignal Lib "GEOCOM32.DLL" (out1 As TMC_EDM_SIGNAL) As Integer
Declare Function VB_TMC_GetPrismCorr Lib "GEOCOM32.DLL" (out1 As Double) As Integer
Declare Function VB_TMC_SetPrismCorr Lib "GEOCOM32.DLL" (in1 As Double) As Integer
Declare Function VB_TMC_GetFace Lib "GEOCOM32.DLL" (out1 As Long) As Integer
Declare Function VB_TMC_SetAtmCorr Lib "GEOCOM32.DLL" (in1 As TMC_ATMOS_TEMPERATURE) As Integer
Declare Function VB_TMC_GetAtmCorr Lib "GEOCOM32.DLL" (out1 As TMC_ATMOS_TEMPERATURE) As Integer
Declare Function VB_TMC_SetRefractiveCorr Lib "GEOCOM32.DLL" (in1 As TMC_REFRACTION) As Integer
Declare Function VB_TMC_GetRefractiveCorr Lib "GEOCOM32.DLL" (out1 As TMC_REFRACTION) As Integer
Declare Function VB_TMC_GetCoordinate1 Lib "GEOCOM32.DLL" (in1 As Long, out2 As TMC_COORDINATE, in3 As Long) As Integer
Declare Function VB_TMC_SetRefractiveMethod Lib "GEOCOM32.DLL" (in1 As Integer) As Integer
Declare Function VB_TMC_GetRefractiveMethod Lib "GEOCOM32.DLL" (out1 As Integer) As Integer
Declare Function VB_TMC_GetAngle5 Lib "GEOCOM32.DLL" (out1 As TMC_HZ_V_ANG, in2 As Long) As Integer
Declare Function VB_TMC_GetSimpleMea Lib "GEOCOM32.DLL" (in1 As Long, out2 As TMC_HZ_V_ANG, out3 As Double, in4 As Long) As Integer
Declare Function VB_TMC_SetOrientation Lib "GEOCOM32.DLL" (in1 As Double) As Integer
Declare Function VB_TMC_IfDataAzeCorrError Lib "GEOCOM32.DLL" (out1 As Long) As Integer
Declare Function VB_TMC_IfDataIncCorrError Lib "GEOCOM32.DLL" (out1 As Long) As Integer
Declare Function VB_TMC_GetSimpleCoord Lib "GEOCOM32.DLL" (in1 As Long, out2 As Double, out3 As Double, out4 As Double, in5 As Long) As Integer
Declare Function VB_TMC_QuickDist Lib "GEOCOM32.DLL" (out1 As TMC_HZ_V_ANG, out2 As Double) As Integer
Declare Function VB_CSV_GetInstrumentNo Lib "GEOCOM32.DLL" (out1 As Long) As Integer
Declare Function VB_CSV_GetInstrumentName Lib "GEOCOM32.DLL" (ByVal out1 As String) As Integer
Declare Function VB_CSV_SetUserInstrumentName Lib "GEOCOM32.DLL" (ByVal in1 As String) As Integer
Declare Function VB_CSV_GetUserInstrumentName Lib "GEOCOM32.DLL" (ByVal out1 As String) As Integer
Declare Function VB_CSV_SetDateTime Lib "GEOCOM32.DLL" (in1 As DATIME) As Integer
Declare Function VB_CSV_GetDateTime Lib "GEOCOM32.DLL" (out1 As DATIME) As Integer
Declare Function VB_CSV_GetVBat Lib "GEOCOM32.DLL" (out1 As Double) As Integer
Declare Function VB_CSV_GetVMem Lib "GEOCOM32.DLL" (out1 As Double) As Integer
Declare Function VB_CSV_GetIntTemp Lib "GEOCOM32.DLL" (out1 As Double) As Integer
Declare Function VB_CSV_GetSWVersion2 Lib "GEOCOM32.DLL" (out1 As Integer, out2 As Integer, out3 As Integer) As Integer
Declare Function VB_CSV_GetDeviceConfig Lib "GEOCOM32.DLL" (out1 As TPS_DEVICE) As Integer
Declare Function VB_MOT_StartController Lib "GEOCOM32.DLL" (in1 As Long) As Integer
Declare Function VB_MOT_StopController Lib "GEOCOM32.DLL" (in1 As Long) As Integer
Declare Function VB_MOT_SetVelocity Lib "GEOCOM32.DLL" (in1 As MOT_COM_PAIR) As Integer
Declare Function VB_MOT_ReadLockStatus Lib "GEOCOM32.DLL" (out1 As Long) As Integer
Declare Function VB_WIR_GetRecFormat Lib "GEOCOM32.DLL" (out1 As Integer) As Integer
Declare Function VB_WIR_SetRecFormat Lib "GEOCOM32.DLL" (in1 As Integer) As Integer
Declare Function VB_AUT_SetTol Lib "GEOCOM32.DLL" (in1 As AUT_POSTOL) As Integer
Declare Function VB_AUT_ReadTol Lib "GEOCOM32.DLL" (out1 As AUT_POSTOL) As Integer
Declare Function VB_AUT_SetTimeout Lib "GEOCOM32.DLL" (in1 As AUT_TIMEOUT) As Integer
Declare Function VB_AUT_ReadTimeout Lib "GEOCOM32.DLL" (out1 As AUT_TIMEOUT) As Integer
Declare Function VB_AUT_LockIn Lib "GEOCOM32.DLL" () As Integer
Declare Function VB_AUT_SetATRStatus Lib "GEOCOM32.DLL" (in1 As Long) As Integer
Declare Function VB_AUT_GetATRStatus Lib "GEOCOM32.DLL" (out1 As Long) As Integer
Declare Function VB_AUT_SetLockStatus Lib "GEOCOM32.DLL" (in1 As Long) As Integer
Declare Function VB_AUT_GetLockStatus Lib "GEOCOM32.DLL" (out1 As Long) As Integer
Declare Function VB_AUT_MakePositioning4 Lib "GEOCOM32.DLL" (in1 As Double, in2 As Double, in3 As Long, in4 As Long, in5 As Long) As Integer
Declare Function VB_AUT_ChangeFace4 Lib "GEOCOM32.DLL" (in1 As Long, in2 As Long, in3 As Long) As Integer
Declare Function VB_AUT_Search2 Lib "GEOCOM32.DLL" (in1 As Double, in2 As Double, in3 As Long) As Integer
Declare Function VB_AUT_GetFineAdjustMode Lib "GEOCOM32.DLL" (out1 As Long) As Integer
Declare Function VB_AUT_SetFineAdjustMode Lib "GEOCOM32.DLL" (in1 As Long) As Integer
Declare Function VB_AUT_FineAdjust3 Lib "GEOCOM32.DLL" (in1 As Double, in2 As Double, in3 As Long) As Integer
Declare Function VB_BMM_BeepOn Lib "GEOCOM32.DLL" (in1 As Integer, in2 As Integer) As Integer
Declare Function VB_BMM_BeepOff Lib "GEOCOM32.DLL" () As Integer
Declare Function VB_BMM_BeepNormal Lib "GEOCOM32.DLL" () As Integer
Declare Function VB_BMM_BeepAlarm Lib "GEOCOM32.DLL" () As Integer
Declare Function VB_CTL_GetUpCounter Lib "GEOCOM32.DLL" (out1 As Integer, out2 As Integer) As Integer
Declare Function VB_SUP_GetConfig Lib "GEOCOM32.DLL" (out1 As Long, out2 As Long, out3 As Long) As Integer
Declare Function VB_SUP_SetConfig Lib "GEOCOM32.DLL" (in1 As Long, in2 As Long, in3 As Long) As Integer
Declare Function VB_SUP_SwitchLowTempControl Lib "GEOCOM32.DLL" (in1 As Long) As Integer
Declare Function VB_BAP_GetLastDisplayedError Lib "GEOCOM32.DLL" (out1 As Integer, out2 As Integer) As Integer
Declare Function VB_BAP_MeasDistanceAngle Lib "GEOCOM32.DLL" (inout1 As Long, out2 As Double, out3 As Double, out4 As Double) As Integer

#End If
