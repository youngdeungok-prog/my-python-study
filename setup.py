from setuptools import setup, find_packages

setup(
    name="ko-schedule-gantt-chart",         # pip install 할 때 쓰는 이름
    version="0.1.1",                         # 버전 (수정할 때마다 올리면 좋습니다)
    author="Yeong-deung",                    # 제작자 이름
    description="A specialized Excel Gantt chart generator for Korean SCM scheduling.",
    long_description=open('README.md', encoding='utf-8').read(),
    long_description_content_type='text/markdown',
    url="https://github.com/youngdeungok-prog/ko-schedule-gantt-chart", # 레파지토리 주소
    packages=find_packages(),                # ko_schedule_gantt 폴더를 자동으로 찾습니다
    install_requires=[                       # 설치 시 자동으로 함께 설치될 라이브러리
        "pandas",
        "xlsxwriter",
    ],
    classifiers=[
        "Programming Language :: Python :: 3",
        "License :: OSI Approved :: MIT License",
        "Operating System :: OS Independent",
    ],
    python_requires='>=3.8',
)