--
-- PostgreSQL database dump
--

\restrict 2dD5E7SnkOObSocnaWAenHS2xiAgyP5PYfK8nUvutZgJoRbDgjrUKvS4E3o3N1Y

-- Dumped from database version 15.17 (Debian 15.17-1.pgdg13+1)
-- Dumped by pg_dump version 15.17 (Debian 15.17-1.pgdg13+1)

SET statement_timeout = 0;
SET lock_timeout = 0;
SET idle_in_transaction_session_timeout = 0;
SET client_encoding = 'UTF8';
SET standard_conforming_strings = on;
-- SELECT pg_catalog.set_config('search_path', '', false);
SET search_path TO public;
SET check_function_bodies = false;
SET xmloption = content;
SET client_min_messages = warning;
SET row_security = off;

SET default_tablespace = '';

SET default_table_access_method = heap;

--
-- Name: chitieu; Type: TABLE; Schema: public; Owner: postgres
--

CREATE TABLE public.chitieu (
    id integer NOT NULL,
    stt text,
    ten_chitieu text,
    cap_dong integer
);


ALTER TABLE public.chitieu OWNER TO postgres;

--
-- Name: chitieu_id_seq; Type: SEQUENCE; Schema: public; Owner: postgres
--

CREATE SEQUENCE public.chitieu_id_seq
    AS integer
    START WITH 1
    INCREMENT BY 1
    NO MINVALUE
    NO MAXVALUE
    CACHE 1;


ALTER TABLE public.chitieu_id_seq OWNER TO postgres;

--
-- Name: chitieu_id_seq; Type: SEQUENCE OWNED BY; Schema: public; Owner: postgres
--

ALTER SEQUENCE public.chitieu_id_seq OWNED BY public.chitieu.id;


--
-- Name: dulieu_baocao; Type: TABLE; Schema: public; Owner: postgres
--

CREATE TABLE public.dulieu_baocao (
    id integer NOT NULL,
    xa_id integer,
    thang integer,
    nam integer,
    chitieu_id integer,
    tiepnhan_tructuyen integer,
    tiepnhan_tructiep integer,
    giaiquyet_tructuyen integer,
    giaiquyet_tructiep integer,
    dangxuly_tructuyen integer,
    dangxuly_tructiep integer,
    thanhtoan_tructuyen integer,
    sohoa integer,
    lam_sach integer,
    chua_lam_sach integer,
    ho_so_qua_han integer,
    tra_kq_tructuyen integer,
    trang_thai text
);


ALTER TABLE public.dulieu_baocao OWNER TO postgres;

--
-- Name: dulieu_baocao_id_seq; Type: SEQUENCE; Schema: public; Owner: postgres
--

CREATE SEQUENCE public.dulieu_baocao_id_seq
    AS integer
    START WITH 1
    INCREMENT BY 1
    NO MINVALUE
    NO MAXVALUE
    CACHE 1;


ALTER TABLE public.dulieu_baocao_id_seq OWNER TO postgres;

--
-- Name: dulieu_baocao_id_seq; Type: SEQUENCE OWNED BY; Schema: public; Owner: postgres
--

ALTER SEQUENCE public.dulieu_baocao_id_seq OWNED BY public.dulieu_baocao.id;


--
-- Name: ky_baocao; Type: TABLE; Schema: public; Owner: postgres
--

CREATE TABLE public.ky_baocao (
    id integer NOT NULL,
    thang integer,
    nam integer,
    trang_thai text DEFAULT 'mo'::text
);


ALTER TABLE public.ky_baocao OWNER TO postgres;

--
-- Name: ky_baocao_id_seq; Type: SEQUENCE; Schema: public; Owner: postgres
--

CREATE SEQUENCE public.ky_baocao_id_seq
    AS integer
    START WITH 1
    INCREMENT BY 1
    NO MINVALUE
    NO MAXVALUE
    CACHE 1;


ALTER TABLE public.ky_baocao_id_seq OWNER TO postgres;

--
-- Name: ky_baocao_id_seq; Type: SEQUENCE OWNED BY; Schema: public; Owner: postgres
--

ALTER SEQUENCE public.ky_baocao_id_seq OWNED BY public.ky_baocao.id;


--
-- Name: users; Type: TABLE; Schema: public; Owner: postgres
--

CREATE TABLE public.users (
    id integer NOT NULL,
    username text,
    password text,
    role text,
    xa_id integer
);


ALTER TABLE public.users OWNER TO postgres;

--
-- Name: users_id_seq; Type: SEQUENCE; Schema: public; Owner: postgres
--

CREATE SEQUENCE public.users_id_seq
    AS integer
    START WITH 1
    INCREMENT BY 1
    NO MINVALUE
    NO MAXVALUE
    CACHE 1;


ALTER TABLE public.users_id_seq OWNER TO postgres;

--
-- Name: users_id_seq; Type: SEQUENCE OWNED BY; Schema: public; Owner: postgres
--

ALTER SEQUENCE public.users_id_seq OWNED BY public.users.id;


--
-- Name: xa; Type: TABLE; Schema: public; Owner: postgres
--

CREATE TABLE public.xa (
    id integer NOT NULL,
    ma_xa text,
    ten_xa text
);


ALTER TABLE public.xa OWNER TO postgres;

--
-- Name: xa_id_seq; Type: SEQUENCE; Schema: public; Owner: postgres
--

CREATE SEQUENCE public.xa_id_seq
    AS integer
    START WITH 1
    INCREMENT BY 1
    NO MINVALUE
    NO MAXVALUE
    CACHE 1;


ALTER TABLE public.xa_id_seq OWNER TO postgres;

--
-- Name: xa_id_seq; Type: SEQUENCE OWNED BY; Schema: public; Owner: postgres
--

ALTER SEQUENCE public.xa_id_seq OWNED BY public.xa.id;


--
-- Name: chitieu id; Type: DEFAULT; Schema: public; Owner: postgres
--

ALTER TABLE ONLY public.chitieu ALTER COLUMN id SET DEFAULT nextval('public.chitieu_id_seq'::regclass);


--
-- Name: dulieu_baocao id; Type: DEFAULT; Schema: public; Owner: postgres
--

ALTER TABLE ONLY public.dulieu_baocao ALTER COLUMN id SET DEFAULT nextval('public.dulieu_baocao_id_seq'::regclass);


--
-- Name: ky_baocao id; Type: DEFAULT; Schema: public; Owner: postgres
--

ALTER TABLE ONLY public.ky_baocao ALTER COLUMN id SET DEFAULT nextval('public.ky_baocao_id_seq'::regclass);


--
-- Name: users id; Type: DEFAULT; Schema: public; Owner: postgres
--

ALTER TABLE ONLY public.users ALTER COLUMN id SET DEFAULT nextval('public.users_id_seq'::regclass);


--
-- Name: xa id; Type: DEFAULT; Schema: public; Owner: postgres
--

ALTER TABLE ONLY public.xa ALTER COLUMN id SET DEFAULT nextval('public.xa_id_seq'::regclass);


--
-- Data for Name: chitieu; Type: TABLE DATA; Schema: public; Owner: postgres
--

COPY public.chitieu (id, stt, ten_chitieu, cap_dong) FROM stdin;
1	A	DỊCH VỤ CÔNG TOÀN TRÌNH	0
2	I	Lĩnh vực Quản lý xuất nhập cảnh	1
3	1	Cấp hộ chiếu phổ thông lần đầu hoặc lần thứ hai trở đi khi hộ chiếu cũ không còn giá trị	2
4	2	Khai báo tạm trú cho người nước ngoài	2
5	3	Trình báo mất giấy thông hành	2
6	4	Trình báo mất hộ chiếu phổ thông	2
7	5	Trình báo mất thẻ ABTC	2
8	II	Lĩnh vực Đăng ký, quản lý cư trú	1
9	6	Điều chỉnh thông tin về cư trú 	2
10	7	Khai báo tạm vắng	2
11	8	Thông báo lưu trú	2
12	9	Xác nhận thông tin về cư trú	2
13	10	Xóa đăng ký tạm trú	2
14	11	Xóa đăng ký thường trú	2
15	III	Lĩnh vực Cấp, quản lý căn cước	1
16	12	Cấp lại thẻ căn cước	2
17	13	Xác nhận số chứng minh nhân dân, số định danh	2
18	IV	Lĩnh vực quản lý ngành nghề đầu tư, kinh doanh có điều kiện về an ninh, trật tự	1
19	14	Cấp mới giấy chứng nhận đủ điều kiện về an ninh, trật tự	2
20	15	Cấp đổi giấy chứng nhận đủ điều kiện về an ninh, trật tự	2
21	16	Cấp lại giấy chứng nhận đủ điều kiện về an ninh, trật tự	2
22	V	Lĩnh vực Đăng ký, quản lý con dấu	1
23	17	Đăng ký dấu nổi, dấu thu nhỏ, dấu xi	2
24	18	Đăng ký mẫu con dấu mới	2
25	19	Đăng ký thêm con dấu	2
26	VI	Lĩnh vực Quản lý vũ khí, vật liệu nổ, công cụ hỗ trợ và pháo	1
27	20	Cấp giấy phép sửa chữa công cụ hỗ trợ	2
28	21	Cấp giấy phép trang bị công cụ hỗ trợ	2
29	22	Cấp giấy phép trang bị vũ khí thể thao	2
30	23	Cấp giấy phép trang bị vũ khí thô sơ	2
31	24	Cấp giấy phép trang bị, giấy phép sử dụng vũ khí, công cụ hỗ trợ để triển lãm, trưng bày hoặc làm đạo cụ trong hoạt động văn hóa, nghệ thuật	2
32	25	Cấp giấy phép vận chuyển công cụ hỗ trợ	2
33	26	Cấp giấy phép vận chuyển pháo hoa để kinh doanh	2
34	27	Cấp giấy phép vận chuyển tiền chất thuốc nổ	2
35	28	Cấp giấy phép vận chuyển vật liệu nổ công nghiệp	2
36	29	Đề nghị đào tạo, huấn luyện về quản lý, sử dụng vũ khí, vật liệu nổ quân dụng, công cụ hỗ trợ và cấp, cấp đổi giấy chứng nhận sử dụng vũ khí, vật liệu nổ quân dụng, công cụ hỗ trợ; chứng chỉ quản lý kho, nơi cất giữ vũ khí, vật liệu nổ quân dụng, công cụ hỗ trợ	2
37	30	Điều chỉnh giấy phép vận chuyển tiền chất thuốc nổ	2
38	31	Điều chỉnh giấy phép vận chuyển vật liệu nổ công nghiệp	2
39	VII	Lĩnh vực Phòng cháy, chữa cháy và cứu nạn, cứu hộ	1
40	32	Cấp Giấy chứng nhận kiểm định phương tiện phòng cháy và chữa cháy (Đang có hiệu lực)	2
41	33	Cấp Giấy phép lưu thông phương tiện phòng cháy, chữa cháy, cứu nạn, cứu hộ  (Đang có hiệu lực)	2
42	34	Thẩm định thiết kế về phòng cháy chữa cháy  (Đang có hiệu lực)	2
43	35	Thẩm duyệt thiết kế về phòng cháy và chữa cháy (trường hợp điều chỉnh thiết kế trong quá trình thi công đối với dự án đầu tư xây dựng công trình, công trình, phương tiện giao thông đã được cơ quan Công an cấp Giấy chứng nhận thẩm duyệt thiết kế về phòng cháy và chữa cháy mà chưa được chấp thuận kết quả nghiệm thu)	2
44	36	Cấp chứng chỉ hành nghề tư vấn về phòng cháy và chữa cháy	2
45	37	Cấp đổi chứng chỉ hành nghề tư vấn về phòng cháy và chữa cháy	2
46	38	Cấp đổi giấy xác nhận đủ điều kiện kinh doanh dịch vụ phòng cháy và chữa cháy	2
47	39	Cấp Giấy chứng nhận kiểm định phương tiện phòng cháy và chữa cháy (Đang có hiệu lực)	2
48	40	Cấp giấy xác nhận đủ điều kiện kinh doanh dịch vụ phòng cháy và chữa cháy	2
49	41	Cấp lại chứng chỉ hành nghề tư vấn về phòng cháy và chữa cháy	2
50	42	Cấp lại giấy phép vận chuyển hàng hóa nguy hiểm về cháy, nổ thuộc loại 1, loại 2, loại 3, loại 4 và loại 9 bằng phương tiện giao thông cơ giới đường bộ, trên đường thủy nội địa (trừ vật liệu nổ công nghiệp)	2
51	43	Cấp lại giấy xác nhận đủ điều kiện kinh doanh dịch vụ phòng cháy và chữa cháy	2
52	44	Nộp phạt vi phạm hành chính trong lĩnh vực phòng cháy, chữa cháy và cứu nạn, cứu hộ	2
53	45	Phê duyệt phương án chữa cháy cơ sở	2
54	46	Thẩm duyệt thiết kế về phòng cháy và chữa cháy	2
55	VIII	Lĩnh vực Giao thông	1
56	47	Cấp giấy phép sử dụng thiết bị phát tín hiệu của xe được quyền ưu tiên	2
57	48	Đăng ký, cấp biển số xe lần đầu đối với xe sản xuất, lắp ráp trong nước	2
58	49	Đăng ký, cấp biển số xe lần đầu đối với xe nhập khẩu	2
59	50	Cấp lại chứng nhận đăng ký xe, biển số xe	2
60	51	Đăng ký xe tạm thời	2
61	52	Thu hồi chứng nhận đăng ký xe, biển số xe	2
62	53	Nộp phạt vi phạm hành chính trong lĩnh vực giao thông	2
63	IX	Lĩnh vực Lý lịch tư pháp	1
64	54	Cấp Phiếu lý lịch tư pháp cho công dân Việt Nam, người nước ngoài cư trú tại Việt Nam	2
65	55	Cấp Phiếu lý lịch tư pháp theo yêu cầu của cơ quan nhà nước, tổ chức chính trị, tổ chức chính trị-xã hội	2
66	56	Cấp Phiếu lý lịch tư pháp theo yêu cầu của cơ quan tiến hành tố tụng	2
67	B	DỊCH VỤ CÔNG MỘT PHẦN	0
68	I	Lĩnh vực Quản lý xuất nhập cảnh	1
69	57	Cấp giấy thông hành biên giới Việt Nam - Campuchia cho cán bộ, công chức, viên chức, công nhân sang Campuchia tại Công an cấp tỉnh biên giới tiếp giáp với Campuchia	2
70	58	Cấp giấy thông hành biên giới Việt Nam - Lào cho công dân Việt Nam làm việc trong các cơ quan, tổ chức, doanh nghiệp có trụ sở tại tỉnh có chung đường biên giới với Lào	2
71	59	Cấp giấy thông hành biên giới Việt Nam - Lào cho công dân Việt Nam có hộ khẩu thường trú ở tỉnh có chung đường biên giới với Lào	2
72	60	Cấp giấy phép xuất nhập cảnh cho người không quốc tịch cư trú tại Việt Nam	2
73	61	Cấp hộ chiếu phổ thông cho người dưới 14 tuổi hoặc từ lần thứ hai trở đi khi hộ chiếu cũ còn giá trị	2
74	62	Cấp lại giấy phép xuất nhập cảnh cho người không quốc tịch cư trú tại Việt Nam	2
75	63	Cấp thẻ tạm trú cho người nước ngoài tại Việt Nam	2
76	64	Cấp thị thực cho người nước ngoài tại Việt Nam	2
77	65	Gia hạn tạm trú cho người nhập cảnh bằng giấy miễn thị thực	2
78	66	Gia hạn tạm trú cho người nước ngoài tại Việt Nam	2
79	67	Khôi phục giá trị sử dụng hộ chiếu phổ thông	2
80	II	Lĩnh vực đảm bảo an ninh hàng không	1
81	68	Cấp mới thẻ kiểm soát an ninh cảng hàng không, sân bay có giá trị sử dụng dài hạn	2
82	69	Cấp lại thẻ kiểm soát an ninh cảng hàng không, sân bay có giá trị sử dụng dài hạn	2
83	70	Cấp thẻ kiểm soát an ninh cảng hàng không, sân bay có giá trị sử dụng ngắn hạn	2
84	71	Cấp mới giấy phép kiểm soát an ninh cảng hàng không, sân bay có giá trị sử dụng dài hạn	2
85	72	Cấp lại giấy phép kiểm soát an ninh cảng hàng không, sân bay có giá trị sử dụng dài hạn	2
86	73	Cấp giấy phép kiểm soát an ninh cảng hàng không, sân bay có giá trị sử dụng ngắn hạn	2
87	III	Lĩnh vực Đăng ký, quản lý cư trú	1
88	74	Đăng ký tạm trú	2
89	75	Đăng ký thường trú	2
90	76	Gia hạn tạm trú	2
91	77	Khai báo thông tin về cư trú 	2
92	78	Tách hộ	2
93	IV	Lĩnh vực Cấp, quản lý căn cước	1
94	79	Cấp đổi thẻ căn cước	2
95	80	Cấp thẻ căn cước cho người dưới 14 tuổi	2
96	81	Cấp thẻ căn cước cho người từ đủ 14 tuổi trở lên	2
97	V	Lĩnh vực Đăng ký, quản lý con dấu	1
98	82	Đăng ký lại mẫu con dấu	2
99	83	Đổi, cấp lại giấy chứng nhận đăng ký mẫu con dấu	2
100	VI	Lĩnh vực Quản lý vũ khí, vật liệu nổ, công cụ hỗ trợ và pháo	1
101	84	Cấp đổi giấy phép sử dụng công cụ hỗ trợ	2
102	85	Cấp đổi giấy phép sử dụng vũ khí thể thao	2
103	86	Cấp giấy phép sử dụng công cụ hỗ trợ	2
104	87	Cấp giấy phép sử dụng vũ khí thể thao	2
105	88	Cấp giấy xác nhận đăng ký công cụ hỗ trợ	2
106	89	Cấp lại giấy phép sử dụng công cụ hỗ trợ	2
107	90	Cấp lại giấy phép sử dụng vũ khí thể thao	2
108	91	Cấp lại giấy xác nhận đăng ký công cụ hỗ trợ	2
109	92	Thông báo khai báo vũ khí thô sơ	2
110	VII	Lĩnh vực Phòng cháy, chữa cháy và cứu nạn, cứu hộ	1
111	93	Kiểm tra công tác nghiệm thu về phòng cháy chữa cháy (Đang có hiệu lực)	2
112	94	Nghiệm thu về phòng cháy chữa cháy (đối với dự án đầu tư xây dựng công trình, công trình, phương tiện giao thông đã được cơ quan Công an cấp Giấy chứng nhận thẩm duyệt thiết kế về phòng cháy và chữa cháy mà chưa được chấp thuận kết quả nghiệm thu) (Đang có hiệu lực)	2
113	95	Phục hồi hoạt động của cơ sở, phương tiện giao thông cơ giới, hộ gia đình và cá nhân (Đang có hiệu lực)	2
114	96	Cấp chứng nhận huấn luyện nghiệp vụ cứu nạn, cứu hộ	2
115	97	Cấp chứng nhận huấn luyện nghiệp vụ phòng cháy, chữa cháy	2
116	98	Cấp giấy phép vận chuyển hàng hóa nguy hiểm về cháy, nổ thuộc loại 1, loại 2, loại 3, loại 4 và loại 9 bằng phương tiện giao thông cơ giới đường bộ, trên đường thủy nội địa	2
117	99	Cấp giấy phép vận chuyển hàng hóa nguy hiểm về cháy, nổ trên đường sắt	2
118	100	Nghiệm thu về phòng cháy và chữa cháy	2
119	101	Phục hồi hoạt động của cơ sở, phương tiện giao thông cơ giới, hộ gia đình và cá nhân	2
120	VIII	Lĩnh vực Giao thông	1
121	102	Đăng ký, cấp biển số xe lần đầu	2
122	103	Đăng ký sang tên, di chuyển xe	2
123	104	Cấp đổi chứng nhận đăng ký xe, biển số xe	2
124	105	Cấp lại chứng nhận đăng ký xe, biển số xe	2
125	106	Đăng ký xe tạm thời	2
126	107	Thu hồi giấy chứng nhận đăng ký xe, biển số xe	2
127	108	Chấp thuận hoạt động của sân tập lái để sát hạch lái xe mô tô 	2
128	109	Chấp thuận lại hoạt động của sân tập lái để sát hạch lái xe mô tô	2
129	110	Thu hồi hoạt động của sân tập lái để sát hạch lái xe mô tô	2
130	111	Cấp giấy phép sát hạch cho trung tâm sát hạch lái xe loại 3	2
131	112	Cấp lại giấy phép sát hạch cho trung tâm sát hạch lái xe loại 3	2
132	113	Thu hồi phép sát hạch cho trung tâm sát hạch lái xe loại 3	2
\.


--
-- Data for Name: dulieu_baocao; Type: TABLE DATA; Schema: public; Owner: postgres
--

COPY public.dulieu_baocao (id, xa_id, thang, nam, chitieu_id, tiepnhan_tructuyen, tiepnhan_tructiep, giaiquyet_tructuyen, giaiquyet_tructiep, dangxuly_tructuyen, dangxuly_tructiep, thanhtoan_tructuyen, sohoa, lam_sach, chua_lam_sach, ho_so_qua_han, tra_kq_tructuyen, trang_thai) FROM stdin;
\.


--
-- Data for Name: ky_baocao; Type: TABLE DATA; Schema: public; Owner: postgres
--

COPY public.ky_baocao (id, thang, nam, trang_thai) FROM stdin;
1	4	2026	mo
\.


--
-- Data for Name: users; Type: TABLE DATA; Schema: public; Owner: postgres
--

COPY public.users (id, username, password, role, xa_id) FROM stdin;
1	PV06	$2b$10$lcQ6Gxspks7orKVdKpmHCut9Btaof3k87DG37mM/YFl3Qhqxl3nT2	xa	1
2	PC06	$2b$10$lcQ6Gxspks7orKVdKpmHCut9Btaof3k87DG37mM/YFl3Qhqxl3nT2	xa	2
3	PC07	$2b$10$lcQ6Gxspks7orKVdKpmHCut9Btaof3k87DG37mM/YFl3Qhqxl3nT2	xa	3
4	PC08	$2b$10$lcQ6Gxspks7orKVdKpmHCut9Btaof3k87DG37mM/YFl3Qhqxl3nT2	xa	4
5	PA08	$2b$10$lcQ6Gxspks7orKVdKpmHCut9Btaof3k87DG37mM/YFl3Qhqxl3nT2	xa	5
6	AnBinh	$2b$10$lcQ6Gxspks7orKVdKpmHCut9Btaof3k87DG37mM/YFl3Qhqxl3nT2	xa	6
7	AnNghia	$2b$10$lcQ6Gxspks7orKVdKpmHCut9Btaof3k87DG37mM/YFl3Qhqxl3nT2	xa	7
8	AuCo	$2b$10$lcQ6Gxspks7orKVdKpmHCut9Btaof3k87DG37mM/YFl3Qhqxl3nT2	xa	8
9	BanNguyen	$2b$10$lcQ6Gxspks7orKVdKpmHCut9Btaof3k87DG37mM/YFl3Qhqxl3nT2	xa	9
10	BangLuan	$2b$10$lcQ6Gxspks7orKVdKpmHCut9Btaof3k87DG37mM/YFl3Qhqxl3nT2	xa	10
11	BaoLa	$2b$10$lcQ6Gxspks7orKVdKpmHCut9Btaof3k87DG37mM/YFl3Qhqxl3nT2	xa	11
12	BinhNguyen	$2b$10$lcQ6Gxspks7orKVdKpmHCut9Btaof3k87DG37mM/YFl3Qhqxl3nT2	xa	12
13	BinhPhu	$2b$10$lcQ6Gxspks7orKVdKpmHCut9Btaof3k87DG37mM/YFl3Qhqxl3nT2	xa	13
14	BinhTuyen	$2b$10$lcQ6Gxspks7orKVdKpmHCut9Btaof3k87DG37mM/YFl3Qhqxl3nT2	xa	14
15	BinhXuyen	$2b$10$lcQ6Gxspks7orKVdKpmHCut9Btaof3k87DG37mM/YFl3Qhqxl3nT2	xa	15
16	CamKhe	$2b$10$lcQ6Gxspks7orKVdKpmHCut9Btaof3k87DG37mM/YFl3Qhqxl3nT2	xa	16
17	CaoDuong	$2b$10$lcQ6Gxspks7orKVdKpmHCut9Btaof3k87DG37mM/YFl3Qhqxl3nT2	xa	17
18	CaoPhong	$2b$10$lcQ6Gxspks7orKVdKpmHCut9Btaof3k87DG37mM/YFl3Qhqxl3nT2	xa	18
19	CaoSon	$2b$10$lcQ6Gxspks7orKVdKpmHCut9Btaof3k87DG37mM/YFl3Qhqxl3nT2	xa	19
20	ChanMong	$2b$10$lcQ6Gxspks7orKVdKpmHCut9Btaof3k87DG37mM/YFl3Qhqxl3nT2	xa	20
21	ChiDam	$2b$10$lcQ6Gxspks7orKVdKpmHCut9Btaof3k87DG37mM/YFl3Qhqxl3nT2	xa	21
22	ChiTien	$2b$10$lcQ6Gxspks7orKVdKpmHCut9Btaof3k87DG37mM/YFl3Qhqxl3nT2	xa	22
23	CuDong	$2b$10$lcQ6Gxspks7orKVdKpmHCut9Btaof3k87DG37mM/YFl3Qhqxl3nT2	xa	23
24	DaBac	$2b$10$lcQ6Gxspks7orKVdKpmHCut9Btaof3k87DG37mM/YFl3Qhqxl3nT2	xa	24
25	DaiDinh	$2b$10$lcQ6Gxspks7orKVdKpmHCut9Btaof3k87DG37mM/YFl3Qhqxl3nT2	xa	25
26	DaiDong	$2b$10$lcQ6Gxspks7orKVdKpmHCut9Btaof3k87DG37mM/YFl3Qhqxl3nT2	xa	26
27	DanChu	$2b$10$lcQ6Gxspks7orKVdKpmHCut9Btaof3k87DG37mM/YFl3Qhqxl3nT2	xa	27
28	DanThuong	$2b$10$lcQ6Gxspks7orKVdKpmHCut9Btaof3k87DG37mM/YFl3Qhqxl3nT2	xa	28
29	DaoTru	$2b$10$lcQ6Gxspks7orKVdKpmHCut9Btaof3k87DG37mM/YFl3Qhqxl3nT2	xa	29
30	DaoXa	$2b$10$lcQ6Gxspks7orKVdKpmHCut9Btaof3k87DG37mM/YFl3Qhqxl3nT2	xa	30
31	DoanHung	$2b$10$lcQ6Gxspks7orKVdKpmHCut9Btaof3k87DG37mM/YFl3Qhqxl3nT2	xa	31
32	DongLuong	$2b$10$lcQ6Gxspks7orKVdKpmHCut9Btaof3k87DG37mM/YFl3Qhqxl3nT2	xa	32
33	DongThanh	$2b$10$lcQ6Gxspks7orKVdKpmHCut9Btaof3k87DG37mM/YFl3Qhqxl3nT2	xa	33
34	DucNhan	$2b$10$lcQ6Gxspks7orKVdKpmHCut9Btaof3k87DG37mM/YFl3Qhqxl3nT2	xa	34
35	DungTien	$2b$10$lcQ6Gxspks7orKVdKpmHCut9Btaof3k87DG37mM/YFl3Qhqxl3nT2	xa	35
36	HaHoa	$2b$10$lcQ6Gxspks7orKVdKpmHCut9Btaof3k87DG37mM/YFl3Qhqxl3nT2	xa	36
37	HaiLuu	$2b$10$lcQ6Gxspks7orKVdKpmHCut9Btaof3k87DG37mM/YFl3Qhqxl3nT2	xa	37
38	HienLuong	$2b$10$lcQ6Gxspks7orKVdKpmHCut9Btaof3k87DG37mM/YFl3Qhqxl3nT2	xa	38
39	HienQuan	$2b$10$lcQ6Gxspks7orKVdKpmHCut9Btaof3k87DG37mM/YFl3Qhqxl3nT2	xa	39
40	HoaBinh	$2b$10$lcQ6Gxspks7orKVdKpmHCut9Btaof3k87DG37mM/YFl3Qhqxl3nT2	xa	40
41	HoangAn	$2b$10$lcQ6Gxspks7orKVdKpmHCut9Btaof3k87DG37mM/YFl3Qhqxl3nT2	xa	41
42	HoangCuong	$2b$10$lcQ6Gxspks7orKVdKpmHCut9Btaof3k87DG37mM/YFl3Qhqxl3nT2	xa	42
43	HoiThinh	$2b$10$lcQ6Gxspks7orKVdKpmHCut9Btaof3k87DG37mM/YFl3Qhqxl3nT2	xa	43
44	HopKim	$2b$10$lcQ6Gxspks7orKVdKpmHCut9Btaof3k87DG37mM/YFl3Qhqxl3nT2	xa	44
45	HopLy	$2b$10$lcQ6Gxspks7orKVdKpmHCut9Btaof3k87DG37mM/YFl3Qhqxl3nT2	xa	45
46	HungViet	$2b$10$lcQ6Gxspks7orKVdKpmHCut9Btaof3k87DG37mM/YFl3Qhqxl3nT2	xa	46
47	HuongCan	$2b$10$lcQ6Gxspks7orKVdKpmHCut9Btaof3k87DG37mM/YFl3Qhqxl3nT2	xa	47
48	HyCuong	$2b$10$lcQ6Gxspks7orKVdKpmHCut9Btaof3k87DG37mM/YFl3Qhqxl3nT2	xa	48
49	KhaCuu	$2b$10$lcQ6Gxspks7orKVdKpmHCut9Btaof3k87DG37mM/YFl3Qhqxl3nT2	xa	49
50	KimBoi	$2b$10$lcQ6Gxspks7orKVdKpmHCut9Btaof3k87DG37mM/YFl3Qhqxl3nT2	xa	50
51	KySon	$2b$10$lcQ6Gxspks7orKVdKpmHCut9Btaof3k87DG37mM/YFl3Qhqxl3nT2	xa	51
52	LacLuong	$2b$10$lcQ6Gxspks7orKVdKpmHCut9Btaof3k87DG37mM/YFl3Qhqxl3nT2	xa	52
53	LacSon	$2b$10$lcQ6Gxspks7orKVdKpmHCut9Btaof3k87DG37mM/YFl3Qhqxl3nT2	xa	53
54	LacThuy	$2b$10$lcQ6Gxspks7orKVdKpmHCut9Btaof3k87DG37mM/YFl3Qhqxl3nT2	xa	54
55	LaiDong	$2b$10$lcQ6Gxspks7orKVdKpmHCut9Btaof3k87DG37mM/YFl3Qhqxl3nT2	xa	55
56	LamThao	$2b$10$lcQ6Gxspks7orKVdKpmHCut9Btaof3k87DG37mM/YFl3Qhqxl3nT2	xa	56
57	LapThach	$2b$10$lcQ6Gxspks7orKVdKpmHCut9Btaof3k87DG37mM/YFl3Qhqxl3nT2	xa	57
58	LienChau	$2b$10$lcQ6Gxspks7orKVdKpmHCut9Btaof3k87DG37mM/YFl3Qhqxl3nT2	xa	58
59	LienHoa	$2b$10$lcQ6Gxspks7orKVdKpmHCut9Btaof3k87DG37mM/YFl3Qhqxl3nT2	xa	59
60	LienMinh	$2b$10$lcQ6Gxspks7orKVdKpmHCut9Btaof3k87DG37mM/YFl3Qhqxl3nT2	xa	60
61	LienSon	$2b$10$lcQ6Gxspks7orKVdKpmHCut9Btaof3k87DG37mM/YFl3Qhqxl3nT2	xa	61
62	LongCoc	$2b$10$lcQ6Gxspks7orKVdKpmHCut9Btaof3k87DG37mM/YFl3Qhqxl3nT2	xa	62
63	LuongSon	$2b$10$lcQ6Gxspks7orKVdKpmHCut9Btaof3k87DG37mM/YFl3Qhqxl3nT2	xa	63
64	MaiChau	$2b$10$lcQ6Gxspks7orKVdKpmHCut9Btaof3k87DG37mM/YFl3Qhqxl3nT2	xa	64
65	MaiHa	$2b$10$lcQ6Gxspks7orKVdKpmHCut9Btaof3k87DG37mM/YFl3Qhqxl3nT2	xa	65
66	MinhDai	$2b$10$lcQ6Gxspks7orKVdKpmHCut9Btaof3k87DG37mM/YFl3Qhqxl3nT2	xa	66
67	MinhHoa	$2b$10$lcQ6Gxspks7orKVdKpmHCut9Btaof3k87DG37mM/YFl3Qhqxl3nT2	xa	67
68	MuongBi	$2b$10$lcQ6Gxspks7orKVdKpmHCut9Btaof3k87DG37mM/YFl3Qhqxl3nT2	xa	68
69	MuongDong	$2b$10$lcQ6Gxspks7orKVdKpmHCut9Btaof3k87DG37mM/YFl3Qhqxl3nT2	xa	69
70	MuongHoa	$2b$10$lcQ6Gxspks7orKVdKpmHCut9Btaof3k87DG37mM/YFl3Qhqxl3nT2	xa	70
71	MuongThang	$2b$10$lcQ6Gxspks7orKVdKpmHCut9Btaof3k87DG37mM/YFl3Qhqxl3nT2	xa	71
72	MuongVang	$2b$10$lcQ6Gxspks7orKVdKpmHCut9Btaof3k87DG37mM/YFl3Qhqxl3nT2	xa	72
73	NatSon	$2b$10$lcQ6Gxspks7orKVdKpmHCut9Btaof3k87DG37mM/YFl3Qhqxl3nT2	xa	73
74	NgocSon	$2b$10$lcQ6Gxspks7orKVdKpmHCut9Btaof3k87DG37mM/YFl3Qhqxl3nT2	xa	74
75	NguyetDuc	$2b$10$lcQ6Gxspks7orKVdKpmHCut9Btaof3k87DG37mM/YFl3Qhqxl3nT2	xa	75
76	NhanNghia	$2b$10$lcQ6Gxspks7orKVdKpmHCut9Btaof3k87DG37mM/YFl3Qhqxl3nT2	xa	76
77	NongTrang	$2b$10$lcQ6Gxspks7orKVdKpmHCut9Btaof3k87DG37mM/YFl3Qhqxl3nT2	xa	77
78	PaCo	$2b$10$lcQ6Gxspks7orKVdKpmHCut9Btaof3k87DG37mM/YFl3Qhqxl3nT2	xa	78
79	PhongChau	$2b$10$lcQ6Gxspks7orKVdKpmHCut9Btaof3k87DG37mM/YFl3Qhqxl3nT2	xa	79
80	PhuKhe	$2b$10$lcQ6Gxspks7orKVdKpmHCut9Btaof3k87DG37mM/YFl3Qhqxl3nT2	xa	80
81	PhuMy	$2b$10$lcQ6Gxspks7orKVdKpmHCut9Btaof3k87DG37mM/YFl3Qhqxl3nT2	xa	81
82	PhuNinh	$2b$10$lcQ6Gxspks7orKVdKpmHCut9Btaof3k87DG37mM/YFl3Qhqxl3nT2	xa	82
83	PhuTho	$2b$10$lcQ6Gxspks7orKVdKpmHCut9Btaof3k87DG37mM/YFl3Qhqxl3nT2	xa	83
84	PhucYen	$2b$10$lcQ6Gxspks7orKVdKpmHCut9Btaof3k87DG37mM/YFl3Qhqxl3nT2	xa	84
85	PhungNguyen	$2b$10$lcQ6Gxspks7orKVdKpmHCut9Btaof3k87DG37mM/YFl3Qhqxl3nT2	xa	85
86	QuangYen	$2b$10$lcQ6Gxspks7orKVdKpmHCut9Btaof3k87DG37mM/YFl3Qhqxl3nT2	xa	86
87	QuyDuc	$2b$10$lcQ6Gxspks7orKVdKpmHCut9Btaof3k87DG37mM/YFl3Qhqxl3nT2	xa	87
88	QuyetThang	$2b$10$lcQ6Gxspks7orKVdKpmHCut9Btaof3k87DG37mM/YFl3Qhqxl3nT2	xa	88
89	SonDong	$2b$10$lcQ6Gxspks7orKVdKpmHCut9Btaof3k87DG37mM/YFl3Qhqxl3nT2	xa	89
90	SonLuong	$2b$10$lcQ6Gxspks7orKVdKpmHCut9Btaof3k87DG37mM/YFl3Qhqxl3nT2	xa	90
91	SongLo	$2b$10$lcQ6Gxspks7orKVdKpmHCut9Btaof3k87DG37mM/YFl3Qhqxl3nT2	xa	91
92	TamDao	$2b$10$lcQ6Gxspks7orKVdKpmHCut9Btaof3k87DG37mM/YFl3Qhqxl3nT2	xa	92
93	TamDuong	$2b$10$lcQ6Gxspks7orKVdKpmHCut9Btaof3k87DG37mM/YFl3Qhqxl3nT2	xa	93
94	TamDuongBac	$2b$10$lcQ6Gxspks7orKVdKpmHCut9Btaof3k87DG37mM/YFl3Qhqxl3nT2	xa	94
95	TamHong	$2b$10$lcQ6Gxspks7orKVdKpmHCut9Btaof3k87DG37mM/YFl3Qhqxl3nT2	xa	95
96	TamNong	$2b$10$lcQ6Gxspks7orKVdKpmHCut9Btaof3k87DG37mM/YFl3Qhqxl3nT2	xa	96
97	TamSon	$2b$10$lcQ6Gxspks7orKVdKpmHCut9Btaof3k87DG37mM/YFl3Qhqxl3nT2	xa	97
98	TanHoa	$2b$10$lcQ6Gxspks7orKVdKpmHCut9Btaof3k87DG37mM/YFl3Qhqxl3nT2	xa	98
99	TanLac	$2b$10$lcQ6Gxspks7orKVdKpmHCut9Btaof3k87DG37mM/YFl3Qhqxl3nT2	xa	99
100	TanMai	$2b$10$lcQ6Gxspks7orKVdKpmHCut9Btaof3k87DG37mM/YFl3Qhqxl3nT2	xa	100
101	TanPheo	$2b$10$lcQ6Gxspks7orKVdKpmHCut9Btaof3k87DG37mM/YFl3Qhqxl3nT2	xa	101
102	TanSon	$2b$10$lcQ6Gxspks7orKVdKpmHCut9Btaof3k87DG37mM/YFl3Qhqxl3nT2	xa	102
103	TayCoc	$2b$10$lcQ6Gxspks7orKVdKpmHCut9Btaof3k87DG37mM/YFl3Qhqxl3nT2	xa	103
104	TeLo	$2b$10$lcQ6Gxspks7orKVdKpmHCut9Btaof3k87DG37mM/YFl3Qhqxl3nT2	xa	104
105	ThaiHoa	$2b$10$lcQ6Gxspks7orKVdKpmHCut9Btaof3k87DG37mM/YFl3Qhqxl3nT2	xa	105
106	ThanhBa	$2b$10$lcQ6Gxspks7orKVdKpmHCut9Btaof3k87DG37mM/YFl3Qhqxl3nT2	xa	106
107	ThanhMieu	$2b$10$lcQ6Gxspks7orKVdKpmHCut9Btaof3k87DG37mM/YFl3Qhqxl3nT2	xa	107
108	ThanhSon	$2b$10$lcQ6Gxspks7orKVdKpmHCut9Btaof3k87DG37mM/YFl3Qhqxl3nT2	xa	108
109	ThanhThuy	$2b$10$lcQ6Gxspks7orKVdKpmHCut9Btaof3k87DG37mM/YFl3Qhqxl3nT2	xa	109
110	ThinhMinh	$2b$10$lcQ6Gxspks7orKVdKpmHCut9Btaof3k87DG37mM/YFl3Qhqxl3nT2	xa	110
111	ThongNhat	$2b$10$lcQ6Gxspks7orKVdKpmHCut9Btaof3k87DG37mM/YFl3Qhqxl3nT2	xa	111
112	ThoTang	$2b$10$lcQ6Gxspks7orKVdKpmHCut9Btaof3k87DG37mM/YFl3Qhqxl3nT2	xa	112
113	ThoVan	$2b$10$lcQ6Gxspks7orKVdKpmHCut9Btaof3k87DG37mM/YFl3Qhqxl3nT2	xa	113
114	ThuCuc	$2b$10$lcQ6Gxspks7orKVdKpmHCut9Btaof3k87DG37mM/YFl3Qhqxl3nT2	xa	114
115	ThungNai	$2b$10$lcQ6Gxspks7orKVdKpmHCut9Btaof3k87DG37mM/YFl3Qhqxl3nT2	xa	115
116	ThuongCoc	$2b$10$lcQ6Gxspks7orKVdKpmHCut9Btaof3k87DG37mM/YFl3Qhqxl3nT2	xa	116
117	ThuongLong	$2b$10$lcQ6Gxspks7orKVdKpmHCut9Btaof3k87DG37mM/YFl3Qhqxl3nT2	xa	117
118	TienLu	$2b$10$lcQ6Gxspks7orKVdKpmHCut9Btaof3k87DG37mM/YFl3Qhqxl3nT2	xa	118
119	TienLuong	$2b$10$lcQ6Gxspks7orKVdKpmHCut9Btaof3k87DG37mM/YFl3Qhqxl3nT2	xa	119
120	TienPhong	$2b$10$lcQ6Gxspks7orKVdKpmHCut9Btaof3k87DG37mM/YFl3Qhqxl3nT2	xa	120
121	ToanThang	$2b$10$lcQ6Gxspks7orKVdKpmHCut9Btaof3k87DG37mM/YFl3Qhqxl3nT2	xa	121
122	TramThan	$2b$10$lcQ6Gxspks7orKVdKpmHCut9Btaof3k87DG37mM/YFl3Qhqxl3nT2	xa	122
123	TrungSon	$2b$10$lcQ6Gxspks7orKVdKpmHCut9Btaof3k87DG37mM/YFl3Qhqxl3nT2	xa	123
124	TuVu	$2b$10$lcQ6Gxspks7orKVdKpmHCut9Btaof3k87DG37mM/YFl3Qhqxl3nT2	xa	124
125	VanBan	$2b$10$lcQ6Gxspks7orKVdKpmHCut9Btaof3k87DG37mM/YFl3Qhqxl3nT2	xa	125
126	VanLang	$2b$10$lcQ6Gxspks7orKVdKpmHCut9Btaof3k87DG37mM/YFl3Qhqxl3nT2	xa	126
127	VanMieu	$2b$10$lcQ6Gxspks7orKVdKpmHCut9Btaof3k87DG37mM/YFl3Qhqxl3nT2	xa	127
128	VanPhu	$2b$10$lcQ6Gxspks7orKVdKpmHCut9Btaof3k87DG37mM/YFl3Qhqxl3nT2	xa	128
129	VanSon	$2b$10$lcQ6Gxspks7orKVdKpmHCut9Btaof3k87DG37mM/YFl3Qhqxl3nT2	xa	129
130	VanXuan	$2b$10$lcQ6Gxspks7orKVdKpmHCut9Btaof3k87DG37mM/YFl3Qhqxl3nT2	xa	130
131	VietTri	$2b$10$lcQ6Gxspks7orKVdKpmHCut9Btaof3k87DG37mM/YFl3Qhqxl3nT2	xa	131
132	VinhAn	$2b$10$lcQ6Gxspks7orKVdKpmHCut9Btaof3k87DG37mM/YFl3Qhqxl3nT2	xa	132
133	VinhChan	$2b$10$lcQ6Gxspks7orKVdKpmHCut9Btaof3k87DG37mM/YFl3Qhqxl3nT2	xa	133
134	VinhHung	$2b$10$lcQ6Gxspks7orKVdKpmHCut9Btaof3k87DG37mM/YFl3Qhqxl3nT2	xa	134
135	VinhPhu	$2b$10$lcQ6Gxspks7orKVdKpmHCut9Btaof3k87DG37mM/YFl3Qhqxl3nT2	xa	135
136	VinhPhuc	$2b$10$lcQ6Gxspks7orKVdKpmHCut9Btaof3k87DG37mM/YFl3Qhqxl3nT2	xa	136
137	VinhThanh	$2b$10$lcQ6Gxspks7orKVdKpmHCut9Btaof3k87DG37mM/YFl3Qhqxl3nT2	xa	137
138	VinhTuong	$2b$10$lcQ6Gxspks7orKVdKpmHCut9Btaof3k87DG37mM/YFl3Qhqxl3nT2	xa	138
139	VinhYen	$2b$10$lcQ6Gxspks7orKVdKpmHCut9Btaof3k87DG37mM/YFl3Qhqxl3nT2	xa	139
140	VoMieu	$2b$10$lcQ6Gxspks7orKVdKpmHCut9Btaof3k87DG37mM/YFl3Qhqxl3nT2	xa	140
141	XuanDai	$2b$10$lcQ6Gxspks7orKVdKpmHCut9Btaof3k87DG37mM/YFl3Qhqxl3nT2	xa	141
142	XuanHoa	$2b$10$lcQ6Gxspks7orKVdKpmHCut9Btaof3k87DG37mM/YFl3Qhqxl3nT2	xa	142
143	XuanLang	$2b$10$lcQ6Gxspks7orKVdKpmHCut9Btaof3k87DG37mM/YFl3Qhqxl3nT2	xa	143
144	XuanLung	$2b$10$lcQ6Gxspks7orKVdKpmHCut9Btaof3k87DG37mM/YFl3Qhqxl3nT2	xa	144
145	XuanVien	$2b$10$lcQ6Gxspks7orKVdKpmHCut9Btaof3k87DG37mM/YFl3Qhqxl3nT2	xa	145
146	YenKy	$2b$10$lcQ6Gxspks7orKVdKpmHCut9Btaof3k87DG37mM/YFl3Qhqxl3nT2	xa	146
147	YenLac	$2b$10$lcQ6Gxspks7orKVdKpmHCut9Btaof3k87DG37mM/YFl3Qhqxl3nT2	xa	147
148	YenLang	$2b$10$lcQ6Gxspks7orKVdKpmHCut9Btaof3k87DG37mM/YFl3Qhqxl3nT2	xa	148
149	YenLap	$2b$10$lcQ6Gxspks7orKVdKpmHCut9Btaof3k87DG37mM/YFl3Qhqxl3nT2	xa	149
150	YenPhu	$2b$10$lcQ6Gxspks7orKVdKpmHCut9Btaof3k87DG37mM/YFl3Qhqxl3nT2	xa	150
151	YenSon	$2b$10$lcQ6Gxspks7orKVdKpmHCut9Btaof3k87DG37mM/YFl3Qhqxl3nT2	xa	151
152	YenThuy	$2b$10$lcQ6Gxspks7orKVdKpmHCut9Btaof3k87DG37mM/YFl3Qhqxl3nT2	xa	152
153	YenTri	$2b$10$lcQ6Gxspks7orKVdKpmHCut9Btaof3k87DG37mM/YFl3Qhqxl3nT2	xa	153
154	admin	$2b$10$0opavPfuCYQ.kD6uJHMyG..HWiOT85kuC4WSaDZi4OsuWtWsVXlDO	admin	\N
\.


--
-- Data for Name: xa; Type: TABLE DATA; Schema: public; Owner: postgres
--

COPY public.xa (id, ma_xa, ten_xa) FROM stdin;
1	PV06	Phòng PV06
2	PC06	Phòng PC06
3	PC07	Phòng PC07
4	PC08	Phòng PC08
5	PA08	Phòng PA08
6	AnBinh	An Bình
7	AnNghia	An Nghĩa
8	AuCo	Âu Cơ
9	BanNguyen	Bản Nguyên
10	BangLuan	Bằng Luân
11	BaoLa	Bao La
12	BinhNguyen	Bình Nguyên
13	BinhPhu	Bình Phú
14	BinhTuyen	Bình Tuyền
15	BinhXuyen	Bình Xuyên
16	CamKhe	Cẩm Khê
17	CaoDuong	Cao Dương
18	CaoPhong	Cao Phong
19	CaoSon	Cao Sơn
20	ChanMong	Chân Mộng
21	ChiDam	Chí Đám
22	ChiTien	Chí Tiên
23	CuDong	Cự Đồng
24	DaBac	Đà Bắc
25	DaiDinh	Đại Đình
26	DaiDong	Đại Đồng
27	DanChu	Dân Chủ
28	DanThuong	Đan Thượng
29	DaoTru	Đạo Trù
30	DaoXa	Đào Xá
31	DoanHung	Đoan Hùng
32	DongLuong	Đồng Lương
33	DongThanh	Đông Thành
34	DucNhan	Đức Nhàn
35	DungTien	Dũng Tiến
36	HaHoa	Hạ Hòa
37	HaiLuu	Hải Lựu
38	HienLuong	Hiền Lương
39	HienQuan	Hiền Quan
40	HoaBinh	Hòa Bình
41	HoangAn	Hoàng An
42	HoangCuong	Hoàng Cương
43	HoiThinh	Hội Thịnh
44	HopKim	Hợp Kim
45	HopLy	Hợp Lý
46	HungViet	Hùng Việt
47	HuongCan	Hương Cần
48	HyCuong	Hy Cương
49	KhaCuu	Khả Cửu
50	KimBoi	Kim Bôi
51	KySon	Kỳ Sơn
52	LacLuong	Lạc Lương
53	LacSon	Lạc Sơn
54	LacThuy	Lạc Thủy
55	LaiDong	Lai Đồng
56	LamThao	Lâm Thao
57	LapThach	Lập Thạch
58	LienChau	Liên Châu
59	LienHoa	Liên Hòa
60	LienMinh	Liên Minh
61	LienSon	Liên Sơn
62	LongCoc	Long Cốc
63	LuongSon	Lương Sơn
64	MaiChau	Mai Châu
65	MaiHa	Mai Hạ
66	MinhDai	Minh Đài
67	MinhHoa	Minh Hòa
68	MuongBi	Mường Bi
69	MuongDong	Mường Động
70	MuongHoa	Mường Hoa
71	MuongThang	Mường Thàng
72	MuongVang	Mường Vang
73	NatSon	Nật Sơn
74	NgocSon	Ngọc Sơn
75	NguyetDuc	Nguyệt Đức
76	NhanNghia	Nhân Nghĩa
77	NongTrang	Nông Trang
78	PaCo	Pà Cò
79	PhongChau	Phong Châu
80	PhuKhe	Phú Khê
81	PhuMy	Phú Mỹ
82	PhuNinh	Phù Ninh
83	PhuTho	Phú Thọ
84	PhucYen	Phúc Yên
85	PhungNguyen	Phùng Nguyên
86	QuangYen	Quảng Yên
87	QuyDuc	Quy Đức
88	QuyetThang	Quyết Thắng
89	SonDong	Sơn Đông
90	SonLuong	Sơn Lương
91	SongLo	Sông Lô
92	TamDao	Tam Đảo
93	TamDuong	Tam Dương
94	TamDuongBac	Tam Dương Bắc
95	TamHong	Tam Hồng
96	TamNong	Tam Nông
97	TamSon	Tam Sơn
98	TanHoa	Tân Hòa
99	TanLac	Tân Lạc
100	TanMai	Tân Mai
101	TanPheo	Tân Pheo
102	TanSon	Tân Sơn
103	TayCoc	Tây Cốc
104	TeLo	Tề Lỗ
105	ThaiHoa	Thái Hòa
106	ThanhBa	Thanh Ba
107	ThanhMieu	Thanh Miếu
108	ThanhSon	Thanh Sơn
109	ThanhThuy	Thanh Thủy
110	ThinhMinh	Thịnh Minh
111	ThongNhat	Thống Nhất
112	ThoTang	Thổ Tang
113	ThoVan	Thọ Văn
114	ThuCuc	Thu Cúc
115	ThungNai	Thung Nai
116	ThuongCoc	Thượng Cốc
117	ThuongLong	Thượng Long
118	TienLu	Tiên Lữ
119	TienLuong	Tiên Lương
120	TienPhong	Tiền Phong
121	ToanThang	Toàn Thắng
122	TramThan	Trạm Thản
123	TrungSon	Trung Sơn
124	TuVu	Tu Vũ
125	VanBan	Vân Bán
126	VanLang	Văn Lang
127	VanMieu	Văn Miếu
128	VanPhu	Vân Phú
129	VanSon	Vân Sơn
130	VanXuan	Vạn Xuân
131	VietTri	Việt Trì
132	VinhAn	Vĩnh An
133	VinhChan	Vĩnh Chân
134	VinhHung	Vĩnh Hưng
135	VinhPhu	Vĩnh Phú
136	VinhPhuc	Vĩnh Phúc
137	VinhThanh	Vĩnh Thành
138	VinhTuong	Vĩnh Tường
139	VinhYen	Vĩnh Yên
140	VoMieu	Võ Miếu
141	XuanDai	Xuân Đài
142	XuanHoa	Xuân Hòa
143	XuanLang	Xuân Lãng
144	XuanLung	Xuân Lũng
145	XuanVien	Xuân Viên
146	YenKy	Yên Kỳ
147	YenLac	Yên Lạc
148	YenLang	Yên Lãng
149	YenLap	Yên Lập
150	YenPhu	Yên Phú
151	YenSon	Yên Sơn
152	YenThuy	Yên Thủy
153	YenTri	Yên Trị
\.


--
-- Name: chitieu_id_seq; Type: SEQUENCE SET; Schema: public; Owner: postgres
--

SELECT pg_catalog.setval('public.chitieu_id_seq', 132, true);


--
-- Name: dulieu_baocao_id_seq; Type: SEQUENCE SET; Schema: public; Owner: postgres
--

SELECT pg_catalog.setval('public.dulieu_baocao_id_seq', 1, false);


--
-- Name: ky_baocao_id_seq; Type: SEQUENCE SET; Schema: public; Owner: postgres
--

SELECT pg_catalog.setval('public.ky_baocao_id_seq', 1, true);


--
-- Name: users_id_seq; Type: SEQUENCE SET; Schema: public; Owner: postgres
--

SELECT pg_catalog.setval('public.users_id_seq', 154, true);


--
-- Name: xa_id_seq; Type: SEQUENCE SET; Schema: public; Owner: postgres
--

SELECT pg_catalog.setval('public.xa_id_seq', 153, true);


--
-- Name: chitieu chitieu_pkey; Type: CONSTRAINT; Schema: public; Owner: postgres
--

ALTER TABLE ONLY public.chitieu
    ADD CONSTRAINT chitieu_pkey PRIMARY KEY (id);


--
-- Name: dulieu_baocao dulieu_baocao_pkey; Type: CONSTRAINT; Schema: public; Owner: postgres
--

ALTER TABLE ONLY public.dulieu_baocao
    ADD CONSTRAINT dulieu_baocao_pkey PRIMARY KEY (id);


--
-- Name: ky_baocao ky_baocao_pkey; Type: CONSTRAINT; Schema: public; Owner: postgres
--

ALTER TABLE ONLY public.ky_baocao
    ADD CONSTRAINT ky_baocao_pkey PRIMARY KEY (id);


--
-- Name: users users_pkey; Type: CONSTRAINT; Schema: public; Owner: postgres
--

ALTER TABLE ONLY public.users
    ADD CONSTRAINT users_pkey PRIMARY KEY (id);


--
-- Name: xa xa_pkey; Type: CONSTRAINT; Schema: public; Owner: postgres
--

ALTER TABLE ONLY public.xa
    ADD CONSTRAINT xa_pkey PRIMARY KEY (id);


--
-- PostgreSQL database dump complete
--

\unrestrict 2dD5E7SnkOObSocnaWAenHS2xiAgyP5PYfK8nUvutZgJoRbDgjrUKvS4E3o3N1Y

-- ĐẶT LẠI SCHEMA
SET search_path TO public;

-- ================== VANBAN ==================
CREATE TABLE IF NOT EXISTS vanban (
    id SERIAL PRIMARY KEY,
    xa_id INTEGER REFERENCES xa(id),
    filename TEXT,
    original_name TEXT,
    thang INTEGER,
    nam INTEGER,
    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
);

-- ================== LOG ==================
CREATE TABLE IF NOT EXISTS log_hoatdong (
    id SERIAL PRIMARY KEY,
    xa_id INTEGER,
    hanh_dong TEXT,
    thoi_gian TIMESTAMP DEFAULT CURRENT_TIMESTAMP
);