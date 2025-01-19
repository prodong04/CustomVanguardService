# settings.py
app = 'dist/app'  # PyInstaller로 만든 실행 파일 경로
volume_name = "MyApp"
format = 'UDZO'
files = [app]
icon = None  # 아이콘 파일 경로가 있다면 설정

# DMG 외형 설정
background = None  # 배경 이미지 설정 가능
window_rect = ((100, 100), (640, 480))
icon_size = 128
