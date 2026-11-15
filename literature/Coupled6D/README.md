# Литература для связанного расчёта N, Mx, My, T, Qx, Qy

PDF в этой папке — неизменённые копии из архивов авторов и издателей. Копии хранятся локально в рабочей папке проекта. В Git включаются библиография, SHA-256 и воспроизводимые команды получения. Перед внешним распространением копий учитывайте условия источника.

Для иностранных работ авторизованные русские переводы по названиям и авторам не найдены. Трассировка формул и пояснения к ним ведутся по-русски.

| Локальный файл | Библиография и источник | SHA-256 |
| --- | --- | --- |
| `Poliotti-Bairan-2020-B-spline-section.pdf` | M. Poliotti, J. M. Bairán. *B-spline sectional model for general 3D effects in reinforced concrete elements*. Engineering Structures 207, 110200 (2020). [DOI](https://doi.org/10.1016/j.engstruct.2020.110200), [авторский PDF](https://upcommons.upc.edu/bitstreams/6ada7d69-15ee-4a9a-bc85-c32aeaa1ecfd/download). [Запись лицензии CC BY-NC-ND 4.0](https://agris.fao.org/search/en/providers/124988/records/6748869c7625988a371d1307). | `0fc6dde8b392b3444fadb6a1d770348df4b5861dbe8e4b766fbd4782dd06fd83` |
| `Poliotti-Bairan-2019-Evolutive-dilatancy.pdf` | M. Poliotti, J. M. Bairán. *A new concrete plastic-damage model with an evolutive dilatancy parameter*. Engineering Structures 189, 541–549 (2019). [DOI](https://doi.org/10.1016/j.engstruct.2019.03.086), [карточка UPC](https://hdl.handle.net/2117/131799), [авторский PDF](https://upcommons.upc.edu/server/api/core/bitstreams/24ee0506-67b8-4f2b-84ef-9577d362bf21/content). | `55b25e1c0ef995543e8f7414f04be670b112c294a9c69d02c1ece20ec9e0a08d` |
| `Bairan-2005-TINSA-thesis.pdf` | J. M. Bairán García. *A non-linear coupled model for the analysis of reinforced concrete sections under bending, shear, torsion and axial forces*. Диссертация UPC, 2005. [DOI](https://doi.org/10.5821/dissertation-2117-93489), [авторский PDF](https://upcommons.upc.edu/bitstream/handle/2117/93489/TJBG1de1.pdf). | `5b653fd1cd05ddf257c62f3121e3f537a370495b0e3ca75a3004618dca8ae142` |
| `Bairan-Mari-WCEE2004.pdf` | J. M. Bairán García, A. R. Marí Bernat. Доклад 840 на 13-й World Conference on Earthquake Engineering, 2004. Расчёт стержня в зоне постоянной поперечной силы и сравнение с опытом Kani. [PDF сборника](https://www.iitk.ac.in/nicee/wcee/article/13_840.pdf). | `34f738a5261288c59d3efbddc3c0f7d3f19fea92038f3dbe4b7c0486377be290` |
| `Poliotti-2020-Thesis.pdf` | M. Poliotti. *A framework for 3D nonlinear analysis of reinforced concrete frames under general loading*. Диссертация UPC, 2020. [Карточка TDX](https://www.tdx.cat/handle/10803/670036), [авторский PDF](https://upcommons.upc.edu/server/api/core/bitstreams/b2895b6b-992e-4800-b82e-cfb7338ae19b/content). | `ad9c021bc838db51d5ff3935d696eba5f17b82539b1ccaedbddec7020d6ad22d` |
| `Budzynski-2008-Diagonal-cracks-torsion-shear.pdf` | W. Budzyński. *Proposition of a new method for the calculation of diagonal crack widths in reinforced concrete elements subjected to combined torsion and shear*. Budownictwo i Architektura, 2(1), с. 37–64, 2008. Польский текст. [Статья и лицензия](https://ph.pollub.pl/index.php/bia/en/article/view/2311), [PDF издателя](https://ph.pollub.pl/index.php/bia/en/article/download/2311/1955/7492). | `023258acfd9d656baee16679fe753f91592e3aff6fd37141e05e44fd96cf2150` |
| `Travush-et-al-2018-Torsion-Bending-Journal.pdf` | В. И. Травуш, Н. И. Карпенко, Вл. И. Колчунов, С. С. Каприелов, А. И. Демьянов, А. В. Конорев. *Результаты экспериментальных исследований конструкций квадратного и коробчатого сечений из высокопрочного бетона при кручении с изгибом*. Строительство и реконструкция, № 6 (80), с. 32–43, 2018. [Официальный PDF выпуска](https://oreluniver.ru/public/file/archive/Building_and_reconstruction_6%2880%29-2018.pdf). | `240b855aead8f1714927114168ff8c6b6df1358855c0fc081915fb4e132b2fc2` |

Исходная теория TINSA и программа валидации изложены также в статьях Bairán García и Marí Bernat (2006), Computers & Structures 84, [часть 1, с. 2254–2263](https://doi.org/10.1016/j.compstruc.2006.08.036), [часть 2, с. 2264–2276](https://doi.org/10.1016/j.compstruc.2006.08.035). В таблице указана полная авторская диссертация с уравнениями и исходными контрольными задачами.

Для [карточки TBU3](OnsongoTBU3Benchmark.md) использованы также таблицы диссертации T. Yamamoto и сравнительный расчёт A. Kagermanov и P. Ceresa. Архив и издатель доступны для чтения, но загрузка локальных копий из рабочей среды не завершилась. [Диссертация в архиве Канады](https://www.collectionscanada.gc.ca/obj/s4/f2/dsk1/tape8/PQDD_0002/MQ45940.pdf) · [статья издателя](https://onlinelibrary.wiley.com/doi/10.1155/2018/8596970).

## Локальное восстановление PDF

Из корня проекта:

```powershell
curl.exe -fsSL 'https://upcommons.upc.edu/bitstreams/6ada7d69-15ee-4a9a-bc85-c32aeaa1ecfd/download' -o 'literature/Coupled6D/Poliotti-Bairan-2020-B-spline-section.pdf'
curl.exe -fsSL 'https://upcommons.upc.edu/server/api/core/bitstreams/24ee0506-67b8-4f2b-84ef-9577d362bf21/content' -o 'literature/Coupled6D/Poliotti-Bairan-2019-Evolutive-dilatancy.pdf'
curl.exe -fsSL 'https://upcommons.upc.edu/bitstream/handle/2117/93489/TJBG1de1.pdf' -o 'literature/Coupled6D/Bairan-2005-TINSA-thesis.pdf'
curl.exe -fsSL 'https://www.iitk.ac.in/nicee/wcee/article/13_840.pdf' -o 'literature/Coupled6D/Bairan-Mari-WCEE2004.pdf'
curl.exe -fsSL 'https://upcommons.upc.edu/server/api/core/bitstreams/b2895b6b-992e-4800-b82e-cfb7338ae19b/content' -o 'literature/Coupled6D/Poliotti-2020-Thesis.pdf'
curl.exe -fsSL 'https://ph.pollub.pl/index.php/bia/en/article/download/2311/1955/7492' -o 'literature/Coupled6D/Budzynski-2008-Diagonal-cracks-torsion-shear.pdf'
curl.exe -fsSL 'https://oreluniver.ru/public/file/archive/Building_and_reconstruction_6%2880%29-2018.pdf' -o 'literature/Coupled6D/Travush-et-al-2018-Torsion-Bending-Journal.pdf'
Get-ChildItem -LiteralPath 'literature/Coupled6D' -Filter '*.pdf' | Get-FileHash -Algorithm SHA256
```
