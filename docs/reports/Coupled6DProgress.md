# Связанный расчёт шести усилий. Журнал разработки

## База

Исходный коммит `3e7da82d3391a90aab8678054584f259eda30fe7`, ветка разработки `codex/coupled-6d`. При начале рабочее дерево было чистым. В репозитории нет настроенного remote.

Проверена сохранённая `workbook/output/RC_Section_NDM.xlsm` с SHA-256 `5786c56088352a0530bdffeaa7beb6701a211152896ee13df85551428d5b10af`. Сводный исходник `VBA_All_Code.txt` имеет SHA-256 `e27cc2072d63c203003cfdc3bfd64e663f687f6d44513c780ec38d7d51f6a069`. Штатный скрипт открывал отдельную копию в реальном Excel и подтвердил неизменность исходной книги в каждом прогоне.

| Набор VBA | No | Yes |
| --- | ---: | ---: |
| `modTestSectionSolver.RunSectionSolverTests` | 3544 / 0 | 3544 / 0 |
| `modTestRegressionBaseline.RunRegressionBaselineTests` | 39 / 0 | 39 / 0 |
| `modTestSectionInputReader.RunSectionInputReaderTests` | 107 / 0 | 107 / 0 |
| `modTestSectionWorkflow.RunWorkflowSnapshotTests` | 151 / 0 | 151 / 0 |

Формат числа в таблице — прошедшие / упавшие проверки. Полные журналы в `docs/regression/Coupled6D/Baseline-No.txt` и `Baseline-Yes.txt`. Полный ON/OFF, часовой Focused, полный Excel UI и AutoCAD на этом этапе не запускались.

```powershell
$macros = 'modTestSectionSolver.RunSectionSolverTests,modTestRegressionBaseline.RunRegressionBaselineTests,modTestSectionInputReader.RunSectionInputReaderTests,modTestSectionWorkflow.RunWorkflowSnapshotTests'
powershell -ExecutionPolicy Bypass -File tools/build_workbook/Run-Audit02ModeTests.ps1 -SourceWorkbook workbook/output/RC_Section_NDM.xlsm -Mode No -ReportPath docs/regression/Coupled6D/Baseline-No.txt -Macro $macros
powershell -ExecutionPolicy Bypass -File tools/build_workbook/Run-Audit02ModeTests.ps1 -SourceWorkbook workbook/output/RC_Section_NDM.xlsm -Mode Yes -ReportPath docs/regression/Coupled6D/Baseline-Yes.txt -Macro $macros
```

## Контракты исходной книги

- `CSectionDefinitionBuilder`, рецепты `ISectionRecipe` и CAD-импорт формируют `CSectionModel`. Для общей области имеются `CSectionGeometryQuery.ConcreteDomain`, `ConnectedRegions`, `LineIntervals`, `CGeometryRegion` и `CFiberMeshBuilder`. Точка задания нагрузки привязана к центру бетонной сетки через `CSectionModel.ConcreteMeshCentroid`.
- Ввод нагрузки — `CLoadCombinationReader`, единицы и знаки — `CUnitSystem`, перенос — `CSectionLoadState`. При нулевых T, Qx и Qy действующая постановка N, Mx, My и её траектории сохраняются.
- `CBatchSectionCalculator` выполняет роли профиля. `CStateProvider` и `CStateRepository` управляют состояниями, `CStateSolutionRunner` и `CLimitSearchCoordinator` разделяют решение и поиск. `CResultMeta`, `EInternalStatus` и `CResultStatusPolicy` задают контракт внутренних и внешних результатов. Новые причины должны проходить через эту цепочку.
- `modWorkbookCalculation` принимает настройки и сохранённую геометрию. `CBatchResultWriter`, `CCalculationSnapshotStore`, `CIndexedSnapshotFile` и `CResultReadSource` ведут Results, файловый снимок и выборочное чтение. Схема и экспорт используют этот снимок. Геометрическое предварительное состояние NDMSNP06 имеет собственный Kind.
- Сборка использует `Build-Workbook.ps1`, `SettingsCatalog.ps1`, `WorkbookLayout.json` и `HelpOutline.ps1`. Исходную и собираемую книгу проверяют на отдельных копиях.

## Текущая работа

Разбираются отдельно кинематика и внутреннее равновесие, пространственный закон бетона, независимые упругие эталоны и совместимость ввода. Время нового расчёта пока не измерено. История материалов, расчёт предела и новая структура книги пока отсутствуют.

Начальная техническая постановка — локальная депланация и искажение сечения по Bairán и Marí. Сравниваются 2D КЭ и кубические B-сплайны для поля перемещений. В правой части слабого равновесия под действием поперечных сил необходим продольный градиент вектора напряжений площадки сечения. Проверяется точное построение этого градиента и его согласование с касательной материала. Для пространственного бетона рассматривается опубликованный пластико-повреждаемый закон Poliotti и Bairán с эволюцией дилатансии.

Следующий этап — зафиксировать точную дискретную постановку и независимо проверить упругие круг, кольцо, прямоугольник и несимметричную область с отверстием. До расчётного статуса прочности должны быть проверены физические параметры, необратимость материала, физический критерий предела и эталоны разрушения.

- [Bairán, диссертация UPC по TINSA](https://upcommons.upc.edu/entities/publication/16c5533f-c9b2-42eb-90ba-fbb618a33063)
- [Poliotti и Bairán, B-spline sectional model](https://doi.org/10.1016/j.engstruct.2020.110200)
- [Poliotti и Bairán, пластичность с повреждением и меняющейся дилатансией](https://doi.org/10.1016/j.engstruct.2019.03.086)

## Первая проверка численного базиса

`CCubicSplineBasis` реализует одномерные кубические функции и производные для
тензорного базиса по Poliotti и Bairán (2020), разд. 3.2, ур. (14)–(15).
Проверка `modTestCoupledSplineBasis.RunCoupledSplineBasisTests` в реальном
Excel дала 31 PASS, 0 FAIL. Исходный бинарный файл остался прежним.
Протокол: `docs/regression/Coupled6D/SplineBasis.txt`.

Скопировать исходную книгу в `docs/regression/Coupled6D/SplineBasis.xlsm` и
собрать компоненты на копии:

```powershell
$fixture = 'docs/regression/Coupled6D/SplineBasis.xlsm'
Copy-Item -LiteralPath 'workbook/output/RC_Section_NDM.xlsm' -Destination $fixture
& 'tools/build_workbook/Update-PrivateWorkbookCode.ps1' -WorkbookPath $fixture -ModuleNames @('CCubicSplineBasis','modTestCoupledSplineBasis')
powershell -ExecutionPolicy Bypass -File tools/build_workbook/Run-Audit02ModeTests.ps1 -SourceWorkbook $fixture -Mode No -ReportPath docs/regression/Coupled6D/SplineBasis.txt -Macro 'modTestCoupledSplineBasis.RunCoupledSplineBasisTests'
```

Четыре неизменённых авторских PDF проверены по заголовку PDF, числу страниц
и SHA-256. Библиография и команды скачивания находятся в
[реестре литературы](../../literature/Coupled6D/README.md).

## Упругая постановка

Введены квадратура `CCoupledSectionQuadrature` по общей точной области NDM,
решение неспряжённой системы `CCoupledLinearSystem` и `CCoupledElasticState`
для однородного изотропного эталона. Переиспользованы
`CSectionGeometryQuery`, `CSectionNumericalFrame`, `CGeometryRegion` и
`CFiberMeshBuilder`. Три цилиндрических поля перемещений разложены по одному
тензорному базису. Четыре жёсткие моды устраняются интегральными ограничениями,
без добавления материальной жёсткости.

Слабая форма с продольной производной напряжений соответствует Bairán (2005),
с. 69–75, ур. (3.30)–(3.43), и с. 85–88, ур. (4.5)–(4.16). Для данной
упругой задачи четыре производные деформаций определяются по
`[N′, Mx′, My′, T′] = [0, Qy, Qx, 0]`. Самостоятельная проверка интегралов
покрывает все шесть компонент. Эта реализация ещё не содержит общего
пластико-повреждаемого материала, арматуры, закона истории и поиска предела.
Нелинейный алгоритм источника также фиксирует энергетические моды и решает
задачу в ортогональном подпространстве, с. 97–101. Упругая проверка не
устанавливает допустимость упрощения этого шага в нелинейной модели.

На копии, сохранённой после компиляции VBA, `RunCoupledSplineBasisTests`
дал 31 PASS / 0 FAIL, `RunCoupledElasticStateTests` — 10 PASS / 0 FAIL.
Относительная погрешность закручивания круга 0,4023%, кольца 0,2995%.
Взвешенная относительная L2-невязка поля касательных напряжений при сдвиге
круга 0,7974%. Интеграл Qx в протоколе точно 10 000 Н. Относительные
допуски до запуска — 1,5% для кривизн и закручивания, 3,5% для поля.
Протокол `docs/regression/Coupled6D/ElasticState.txt` относится к
тестовой, а не к публикуемой книге.

Для воспроизведения после подготовки SplineBasis.xlsm:

```powershell
& 'tools/build_workbook/Update-PrivateWorkbookCode.ps1' -WorkbookPath 'docs/regression/Coupled6D/SplineBasis.xlsm' -ModuleNames @('CCoupledSectionQuadrature','CCoupledLinearSystem','CCoupledElasticState','modTestCoupledElasticState')
powershell -ExecutionPolicy Bypass -File tools/build_workbook/Run-Audit02ModeTests.ps1 -SourceWorkbook docs/regression/Coupled6D/SplineBasis.xlsm -Mode No -ReportPath docs/regression/Coupled6D/ElasticState.txt -Macro 'modTestCoupledSplineBasis.RunCoupledSplineBasisTests,modTestCoupledElasticState.RunCoupledElasticStateTests'
```

Следующая контрольная точка: независимые упругие сравнения для невыпуклой
асимметричной области с отверстием, совместная задача с ненулевыми шестью
усилиями и подготовка пространственного материала и поперечной арматуры.

## Поперечная арматура и совместная упругая проверка

Подготовка `CTransverseRebarLayout` и `CTransverseRebarQuadrature` использует
завершённый `CGeometryPath` для каждой группы. Группа задаёт диаметр, шаг
вдоль элемента и класс стали. Пустая схема означает отсутствие данных.
Без поперечной арматуры требуется явное подтверждение. Замкнутый путь
и открытая ветвь передаются в расчёт с фактической геометрией.
Проверяется минимальное расстояние всего пути до бетона и отверстий,
включая диаметр стержня. Принято идеальное сцепление на заданном пути.
Численный вес ветви `As dl/s` выражен в мм². При интегрировании материала
тот же объём должен быть вычтен из бетона. Внутренняя проекция использует
`tx² εxx + ty² εyy + tx ty γxy` для инженерного угла сдвига.

В Excel выполнены 9 проверок подготовки арматуры, 18 упругих проверок
и 1319 проверок рабочей геометрии. Совместный упругий случай охватывает
`N = −52000 Н`, `Mx = 3100000 Н·мм`, `My = −1900000 Н·мм`,
`T = 2200000 Н·мм`, `Qx = 4400 Н`, `Qy = −3700 Н` на невыпуклой
области с эксцентричным круглым отверстием. Проверены интегралы всех
усилий, три параметра плоского поля по точным A/I и суперпозиция
напряжений. Протоколы находятся в
`docs/regression/Coupled6D/TransverseRebar.txt` и
`docs/regression/Coupled6D/AsymmetricElastic.txt`. Расчётная проверка
физического вклада хомутов, их ввод в рабочей книге и независимый КЭ-эталон
для произвольного контура остаются следующими этапами.

## Ввод хомутов по контуру и объёмное замещение

Контрольная точка включает `AddContourGroup`. Пользователь задаёт тип схемы,
диаметр, шаг, защитный слой и при необходимости ID выбранной границы.
Общий геометрический запрос берёт аналитический контур сечения с направлением
«бетон слева». Ось хомута строится смещением на защитный слой плюс половина
диаметра. Для отверстия положительное смещение идёт наружу в бетон, для
наружного контура — внутрь. Самопересечения и несовместимое положение
ветвей проверяются общими геометрическими методами. Принятый пользовательский
интерфейс размещается на листе «Сечения» рядом с армированием каждого типа.
В таблице сочетаний остаются нагрузки и настройки сочетания.

В упругом ядре обе группы арматуры включены в те же уравнения, что и бетон.
В каждой точке стержня матрица поправки `Es p⊗p − Dбетона` учитывается
с площадью продольного стержня или весом `As dl/s` поперечной ветви.
При `ν = 0` сопоставлены осевая жёсткость и изгиб от эксцентричного
удалённого объёма бетона с независимыми аналитическими интегралами.
Разность для продольной деформации около 4,75·10⁻⁹ относительно эталона,
для малой кривизны около 0,31%. В Excel пройдены 12 проверок поперечной
арматуры и 21 упругая проверка. Протокол
`docs/regression/Coupled6D/ReinforcedElastic.txt`.

Следующая работа — добавить совместные напряжения и историю материалов
в итерации, разделить принятие шага и отмену, затем проверить нелинейные
сценарии и установить инженерные критерии предельной точки.

## История деформирования арматуры

`CCoupledSteelModel` задаёт одноосную упругопластическую модель с линейным
кинематическим упрочнением, как в Poliotti и Bairán (2020), разд. 4.
Для двух знаков допускаются разные сопротивления, наклоны упрочнения и
пределы деформации. Расчёт пробного состояния возвращает напряжение,
касательный модуль, пластическую деформацию и обратное напряжение по
переданному подтверждённому состоянию. Хранилище и решение сечения будут
отдельно отвечать за публикацию пробного шага. Рабочие диаграммы исходной
трёхкомпонентной постановки не изменены.

`modTestCoupledSteelModel.RunCoupledSteelTests` в Excel дал 8 PASS / 0 FAIL.
Сверены предел текучести, упрочнение, разгрузка, обратная текучесть, повтор
от подтверждённого состояния после отброшенной пробы и два физических
ограничения деформации. Протокол
`docs/regression/Coupled6D/SteelHistory.txt`. Связанный нелинейный расчёт
сечения и проверки на экспериментальных данных ещё не завершены.
