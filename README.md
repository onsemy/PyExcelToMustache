# PyExcelToMustache

단일 `Excel(xlsx) File`에 대하여 `Sheet`별로 `Mustache Template`으로 변환하는 Python Script

## 용도

개인적인 프로젝트에 쓰기 위한 스크립트입니다. 어떠한 환경에서도 돌아가게 하기 위해 만들었습니다.

```Excel```에 미리 정의된 Schema를 ```Mustache Template```에 정해진 Role에 따라 생성해줍니다.

## 구동환경

- `Python 3.6` 이상이 동작하는 모든 OS
- [`virtualenv`](https://virtualenv.pypa.io/en/stable/)로 쉽게 설정 가능

### 의존성 도구들

`pipenv`설정 파일인 `Pipfile`에 의존성 패키지들이 기록되어 있습니다.

- [`openpyxl`](https://openpyxl.readthedocs.io/en/stable/): `Excel`을 불러오는 데 쓰입니다.
- [`pystache (mustache)`](https://github.com/defunkt/pystache): `Mustache`를 다루는 데 쓰입니다.

## 설치

`Python 3.6` 이상 버전을 설치 후, 아래 구문을 실행합니다.

> ```$ pip install -r requirements.txt```

## Excel Sheet 규칙

1. `Sheet` 이름의 맨 앞 글자에 ```_```가 들어가는 경우 Template에 삽입되지 않습니다.

- 예) ```_info```

2. Sheet에서 1~5번행은 항상 지켜져야 함.

> 2번의 경우에는 [`PyExcelToBSON`](https://github.com/onsemy/PyExcelToBSON)에 필요한 규칙이므로 그대로 따라갑니다.

- 1번행: 해당 열에 대한 설명
- 2번행: 해당 열의 용도 (사용되지 않을 예정)
- 3번행: 해당 열의 Attribute
- 4번행: 해당 열의 Data Type
- 5번행: 해당 열의 이름 (변수명)

## 사용법

```$ PyExcelToMustache.py -i sample.xlsx -t template.mustache -o output```

1. ```sample.xlsx```와 같은 형식의 Excel(xlsx) 파일을 준비합니다. (Repository에 있음)

2. 적절한 인자 값을 넣고 ```PyExcelToMustache.py```을 실행합니다. 인자를 넣는 순서는 상관없습니다.

- ```-i```, ```--input```: Excel(xlsx) 경로와 파일 이름. **실행 시 반드시 입력해야 합니다.**
- ```-t```, ```--template```: Mustache Template 파일의 경로. **실행 시 반드시 입력해야 합니다.**
- ```-o```, ```--output```: `output`폴더의 경로를 지정. **실행 시 반드시 입력해야 합니다.**
- ```-c```, ```--clean```: `output`폴더 정리

3. `output`폴더에 나온 결과물을 확인합니다.

## 해야할 일

- 코드 리펙토링
- (가능하면) Google Sheet 연동
