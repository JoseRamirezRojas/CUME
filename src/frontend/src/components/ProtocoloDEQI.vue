<template>

<div class="row">
  <div class="column" style="background-color: #e3f2fd">
    <MDBListGroup  id="inner" >
      <MDBListGroupItem tag="button" @click="navega('/protocolos-info')" action >Protocolos</MDBListGroupItem>
      <MDBListGroupItem tag="button" @click="navega('/protocolo-fisicoquimica')"> Calidad fisicoquímica </MDBListGroupItem>
      <MDBListGroupItem tag="button" @click="navega('/protocolo-hyqi')" action > HYQI </MDBListGroupItem>
      <MDBListGroupItem tag="button" action active >DEQI</MDBListGroupItem>
      <MDBListGroupItem tag="button" @click="navega('/cuenca-vegetacion')" action>MEQI</MDBListGroupItem>
      <MDBListGroupItem tag="button" @click="navega('/cuenca-conservacion')" action>BEMEQI </MDBListGroupItem>
      <MDBListGroupItem tag="button" @click="navega('/cuenca-jurisdiccion')" action >Material de trabajo</MDBListGroupItem>
    </MDBListGroup>

  </div>

  <div class="column-2" style="background-color:#F8F8F8;">
    <MDBTabs v-model="activeTabId">
      <!-- Tabs navs -->
      <MDBTabNav pills justify tabsClasses="mb-3">
        <MDBTabItem tabId="info-deqi" href="info-deqi">Información del protocolo</MDBTabItem>
        <MDBTabItem tabId="calculo-deqi" href="calculo-deqi">Cálculo  de índice DEQI </MDBTabItem>
      </MDBTabNav>
      <!-- Tabs navs -->
      <!-- Tabs content -->
      <MDBTabContent>
        <!-- info protocolo-->
        <MDBTabPane tabId="info-deqi">
          <h1 class="display-6">Diatomeas</h1>
          <p align="justify" > Las diatomeas (<i>Bacillariophyta</i>) son algas unicelulares que se distinguen por poseer una pared celular de sílice (<b>frústulo</b>), la cual se compone de dos mitades que encajan entre sí, llamadas <i>tecas</i>, la superior (<b>epiteca</b>) es mayor y envuelve parcialmente a la inferior (<b>hipoteca</b>).  </p>
          <img src="../assets/explic_diatomeas.jpg"  class="img-fluid  hover-shadow" alt="diatomea" style=" width: 394px;" /> <br> <br>
          <p align="justify" > Cada teca está formada por una <b>valva</b> y un <b> cíngulo</b>, la valva presenta una diversidad de perforaciones y ornamentaciones variables entre las especies que proveen rutas para el intercambio de gases y nutrientes, así como para la secreción de productos celulares que les permiten formar colonias, adherirse a sustratos y desplazarse en ellos. </p>
          <p align="justify" > Las diatomeas se distribuyen en todos los sistemas acuáticos y son productores primarios que responden de manera directa ante los cambios fisicoquímicos, hidromorfológicos y las interacciones biológicas que ocurren en los ríos. Es un grupo biológico que presenta una gran diversidad de especies con distinto grado de sensibilidad o tolerancia ante diferentes condiciones o nivel trófico. </p> 
          <p align = "justify"> Por su particular ciclo de vida corto, permiten reconocer perturbaciones a corto y largo plazo. En América Latina, donde la etapa de exploración de la diversidad de diatomeas aún sigue en desarrollo, los principales estudios utilizando a las diatomeas como bioindicadores se han desarrollado en Argentina, Brasil, México, Colombia y Ecuador. </p>
          <h1 class="display-6">Colecta</h1>
          <p align = "justify"> Para su colecta se debe seleccionar un transecto de 10 metros de longitud del río, preferentemente en zonas bien iluminadas y con flujo laminar del agua, es decir, evitar zonas de rápidos. En el sitio de muestreo se debe favorecer una zona donde predomine sustrato rocoso en el lecho del río. Cabe destacar que no se recomienda su colecta en la temporada de lluvias, pues los ensambles de diatomeas se encuentran bajo perturbación intensa. </p>
          <img src="../assets/colecta_diatomeas.jpg"  class="img-fluid  hover-shadow" alt="diatomea" style=" width: 329px; " /> <br> <br>
          <p align="justify"  >Se recomienda colectar preferentemente rocas/cantos de tamaño mediano a grande que se encuentren en la zona de flujo, que difícilmente puedan ser desplazados por la corriente, pero que se puedan retirar del lecho y que el sustrato se encuentre totalmente sumergido. Palpar el sustrato para confirmar que tenga una textura resbalosa provocada por el biofilm de diatomeas, y evitar preferentemente aquel sustrato que tenga por encima una capa de sedimento fino depositado. El sustrato se cepillará con un cepillo de dientes nuevo para evitar la contaminación entre muestras en un área de 20 cm<sup>2</sup> de cada uno (100 cm<sup>2</sup> en total). Se cepillará desde los extremos hacia el centro de la unidad, enjuagando el cepillo vigorosamente en tres ocasiones dentro del frasco con alcohol al 70% previamente etiquetado.  </p>
          <p align="justify">  La identificación de las diatomeas se realiza con base en las características morfológicas del frústulo, y para ello se requiere oxidar el contenido celular y materia orgánica que las acompaña. Para ello se recomienda el método de oxidación por peróxido de hidrógeno. Para la oxidación se requiere en principio homogeneizar manualmente la muestra colectada en el frasco de plástico y transferir 20 ml de muestra a un vaso de precipitados de 50 ml. </p> 
          <p align="justify"> La muestra se calienta a 120 °C en una parrilla eléctrica hasta que el volumen se reduzca a 10 ml. Se disminuye la temperatura a 80 °C y se añade 20 ml de peróxido al 30%, teniendo cuidado de que no hierva. </p>
          <p align="justify"> El volumen de la muestra se reduce de nuevo a 10 ml y volvemos a añadir 20 ml de peróxido de hidrógeno. Este paso se repite hasta que observemos que se ha oxidado la materia orgánica de la muestra (cuando se observan pocas o nulas burbujas).</p>
          <p align="justify" > Una vez que la muestra se enfría, se añaden 10 gotas de ácido clorhídrico al 10% para eliminar carbonatos, si existieran. Se transfiere la muestra a un tubo de ensayo de 15 ml, y se añade un poco de agua destilada del vaso de precipitados para transferir el remanente de la muestra hasta aforar a 12 ml el tubo de ensayo.</p>
          <p align="justify"> La muestra se deja sedimentar 24 horas en el tubo de ensayo y se retira el sobrenadante hasta dejar 5 ml de muestra. La muestra se resuspende y afora de nuevo a 12 ml con agua destilada, este se considera el primer enjuague.</p>
          <p align="justify"> Se deben realizar entre 4 y 6 enjuagues para retirar el peróxido de hidrógeno y ácido clorhídrico de la muestra. En el último enjuague se debe retirar el sobrenadante hasta dejar menos de 2 ml de muestra, entonces se resuspende la muestra y con una pipeta Pasteur se transfiere al tubo Eppendorf hasta aforar a 2 ml. Esta muestra se considera el material oxidado en concentración 1:1, listo para montar en la preparación permanente. </p>
          <p align="justify"> Para el proceso de montaje en preparaciones permanentes se elige una superficie plana en una zona con poca o nula vibración. Se prepara la resina de inclusión hasta dejarla a punto de miel. Se coloca un cubreobjetos en la superficie elegida y se vierten en este 200 μl de agua destilada con la micropipeta. Se colocan en el centro del cubreobjetos 200 μl de la muestra de diatomeas y se deja evaporar la muestra a temperatura ambiente por 24 horas. Una vez seca se puede retirar el exceso de resina con una navaja, y la guardamos en una caja de preparaciones para ser revisada las veces que sea necesaria.  </p>
          <h1 class="display-6">Análisis</h1>
          <img src="../assets/diatom_ilustrativa.png"  class="img-fluid  hover-shadow" alt="diatomea" style="float:right; height: 267px; margin-left: 20px" />
          <p align="justify" > Las muestras permanentes se revisan en el microscopio óptico con el objetivo de 100X (utilizando aceite de inmersión). Se recomienda realizar previamente un transecto horizontal en la parte media de la preparación para identificar la diversidad de especies presentes en la muestra, utilizando recursos bibliográficos taxonómicos especializados. En particular, la diversidad acuática de los ríos de la cuenca de México se ha descrito por Carmona-Jiménez et al. (2016) y Salinas-Camarillo et al. (2020); o bien, se puede seguir la siguiente bibliografía de apoyo por si llegase a presentar otra especie: <a href="https://www.researchgate.net/publication/318760666_Freshwater_Benthic_Diatoms_of_Central_Europe_Over_800_common_species_used_in_ecological_assessment_English_edition_with_updated_taxonomy_and_added_species" target="_blank" rel="noopener noreferrer" > Cantonati et al. (2017) </a> y  <a href="https://diatoms.org" target="_blank" rel="noopener noreferrer" > diatoms of north america</a>.</p>
          <p align="justify"> La cuantificación de las diatomeas se realiza observando campos visuales de manera aleatoria sin repetir.   En cada campo visual se toma registro del número de valvas o frústulos de cada especie hasta llegar a contar 500 diatomeas tomando en cuenta que se cuantificarán valvas y frústulos completos por igual; sólo se cuantificarán valvas rotas cuando tengan más del 50% completo; sólo se cuantificarán valvas cuando al menos 50% de su superficie esté dentro del campo visual observado. Al finalizar la cuantificación de los 500 individuos se calcula la abundancia relativa porcentual de las especies.</p>
          <h1 class="display-6">Evaluación</h1>
          <p align="justify" >El valor indicador de las especies está basado en el estudio de Salinas-Camarillo et al. (2020), en donde las especies indicadoras de una alta calidad ecológica tienen un valor de 1, aquellas de calidad ecológica buena tienen un valor de 2, las indicadoras de calidad moderada tienen un valor de 3, las de calidad pobre tienen un valor de 4, y las de mala calidad tienen un valor de 5. </p> <br>
          <h3> <MDBBadge color="info" class="text-wrap" >Los valores indicadores de las especies más abundantes y con mayor distribución se pueden consultar en la hoja de cálculo que puede obtenerse en la pestaña 'CÁLCULO DE ÍNDICE DEQI' llenando el formulario del conteo de especies</MDBBadge> </h3> <br>
          <p align="justify"> La riqueza de diatomeas en la cuenca de México excede las 500 especies, por lo que aquellas que no se presenten en la tabla suelen tener una baja abundancia o suelen ser raras, por ello se considerarán con un valor indicador (<i>v</i>) de 1, de esta manera el cálculo del índice DEQI no se ve modificado. El índice DEQI se calcula a partir de la suma de los producto del valor indicador (<i>v)</i> de cada especie por su abundancia relativa porcentual (<i>h</i>), dividida entre la suma de la abundancia relativa porcentual de todas las especies (debería ser 100%). lo anterior se resume en la siguiente fórmula: </p>
          <p align="justify" class="fs-5"> &emsp; &emsp; <i>DEQI = <u> &emsp; ∑(v<sub>i</sub> · h<sub>i</sub> ) &emsp;</u>  </i>  <br> &emsp; &emsp; &emsp;&emsp;&emsp;&emsp; &emsp; <i>∑h<sub>i</sub> </i> </p>
          <p align="justify">Los valores del DEQI varían de 1.0 a 5.0, indicando cinco diferentes niveles de calidad ecológica. sitios con un valor de alta calidad ecológica se pueden considerar como sitios de referencia. La buena calidad ecológica es el objetivo a alcanzar para todos los sistemas acuáticos, pues refleja condiciones óptimas en la estructura y función de los ecosistemas.</p>
          <p align="justify"> Conforme la calidad ecológica de los sitios se deteriora, se alcanza una calidad moderada, reflejando presiones en aspectos fisicoquímicos del agua como incremento de nutrientes, o una alteración en la estructura hidromorfológica del sistema como represamientos o canalizaciones. Los sitios con calidad pobre suelen presentar un incremento en las concentraciones de nutrientes, así como modificaciones en la naturalidad del cauce. Los sitios con mala calidad suelen ser los que presentan condiciones extremas de contaminación dadas por  aumento en las concentraciones de nutrientes así como por el vertido de aguas grises junto con la modificación de la estructura hidromorfológica del cauce.  </p>
          <MDBTable border="gray" bordered>
            <thead >
              <tr>
                <th scope="col">Calidad ecológica</th>
                <th scope="col">DEQI</th>
              </tr>
            </thead>
            <tbody>
              <tr class="table-primary">
                <td > Alta </td>
                <td> 1 - 1.5</td>
              </tr>
              <tr class="table-success">
                <td > Buena </td>
                <td> 1.5 - 2.5</td>
              </tr>
              <tr class="table-warning">
                <td > Moderada </td>
                <td> 2.5 - 3.5</td>
              </tr>
              <tr class="table-danger">
                <td > Pobre </td>
                <td> 3.5 - 4.5 </td>
              </tr>
              <tr class="bg-danger p-2 bg-opacity-25">
                <td > Mala </td>
                <td> 4.5 - 5 </td>
              </tr>
            </tbody>
          </MDBTable>
        </MDBTabPane>

        <!-- Calculo con microservicio-->
        <MDBTabPane tabId="calculo-deqi">
        <h1 class="display-5">Cálculo del protocolo</h1>
        <br> <p align="justify"> Aquí puedes introducir el conteo de especies de diatomeas analizados y obtenidos en la colecta, hasta abajo encontrarás un ejemplo del llenado de esta ficha. Al dar clic en Generar se descargará un archivo editable de Excel donde puedes visualizar los valores indicadores de cada taxa y el índice de calidad ecológica según este bioindicador.</p>
        <MDBRow tag="form" class="needs-validation" @submit.prevent="genera" >
            <MDBRow> 
              <MDBCol >
                <p> <b>Parámetros físicos</b> </p>
              </MDBCol>
            </MDBRow>
          <MDBRow> 
            <MDBCol >
              <p> Temp. agua <i>(°C) </i> </p>
            </MDBCol>
            <MDBCol >
              <MDBInput
                label="test 1"
                v-model="form.temperatura1"
                invalidFeedback="Llena este campo"
                validFeedback="¡Listo!"
                validationEvent="input"
                required
                pattern="^[+-]?([0-9]+\.?[0-9]*|\.[0-9]+)$"
                title="Favor de introducir un número"
              />
            </MDBCol>
            <MDBCol >
              <MDBInput
                label="test 2"
                v-model="form.temperatura2"
                invalidFeedback="Llena este campo"
                validFeedback="¡Listo!"
                validationEvent="input"
                required
                pattern="^[+-]?([0-9]+\.?[0-9]*|\.[0-9]+)$"
                title="Favor de introducir un número"
              />
            </MDBCol>
            <MDBCol >
              <MDBInput
                label="test 3"
                v-model="form.temperatura3"
                invalidFeedback="Llena este campo"
                validFeedback="¡Listo!"
                validationEvent="input"
                required
                pattern="^[+-]?([0-9]+\.?[0-9]*|\.[0-9]+)$"
                title="Favor de introducir un número"
              >
              </MDBInput>
            </MDBCol>
          </MDBRow>
          <MDBRow> 
            <MDBCol >
              <p>Conduct. específica <i> (μS)</i> </p>
            </MDBCol>
            <MDBCol >
              <MDBInput
                label="test 1"
                v-model="form.conductividad1"
                invalidFeedback="Llena este campo"
                validFeedback="¡Listo!"
                validationEvent="input"
                required
                pattern="^[+-]?([0-9]+\.?[0-9]*|\.[0-9]+)$"
                title="Favor de introducir un número"
              />
            </MDBCol>
            <MDBCol >
              <MDBInput
                label="test 2"
                v-model="form.conductividad2"
                invalidFeedback="Llena este campo"
                validFeedback="¡Listo!"
                validationEvent="input"
                required
                pattern="^[+-]?([0-9]+\.?[0-9]*|\.[0-9]+)$"
                title="Favor de introducir un número"
              />
            </MDBCol>
            <MDBCol >
              <MDBInput
                label="test 3"
                v-model="form.conductividad3"
                invalidFeedback="Llena este campo"
                validFeedback="¡Listo!"
                validationEvent="input"
                required
                pattern="^[+-]?([0-9]+\.?[0-9]*|\.[0-9]+)$"
                title="Favor de introducir un número"
              >
              </MDBInput>
            </MDBCol>
          </MDBRow>
          <MDBRow> 
            <MDBCol >
              <p>Oxígeno disuelto <i> (%)</i> </p>
            </MDBCol>
            <MDBCol >
              <MDBInput
                label="test 1"
                v-model="form.oxiDisuelto1"
                invalidFeedback="Llena este campo"
                validFeedback="¡Listo!"
                validationEvent="input"
                required
                pattern="^[+-]?([0-9]+\.?[0-9]*|\.[0-9]+)$"
                title="Favor de introducir un número"
              />
            </MDBCol>
            <MDBCol >
              <MDBInput
                label="test 2"
                v-model="form.oxiDisuelto2"
                invalidFeedback="Llena este campo"
                validFeedback="¡Listo!"
                validationEvent="input"
                required
                pattern="^[+-]?([0-9]+\.?[0-9]*|\.[0-9]+)$"
                title="Favor de introducir un número"
              />
            </MDBCol>
            <MDBCol >
              <MDBInput
                label="test 3"
                v-model="form.oxiDisuelto3"
                invalidFeedback="Llena este campo"
                validFeedback="¡Listo!"
                validationEvent="input"
                required
                pattern="^[+-]?([0-9]+\.?[0-9]*|\.[0-9]+)$"
                title="Favor de introducir un número"
              >
              </MDBInput>
            </MDBCol>
          </MDBRow>
          <MDBRow> 
            <MDBCol >
              <p>Oxígeno solubre <i> (mg/l)</i> </p>
            </MDBCol>
            <MDBCol >
              <MDBInput
                label="test 1"
                v-model="form.oxiSolubre1"
                invalidFeedback="Llena este campo"
                validFeedback="¡Listo!"
                validationEvent="input"
                required
                pattern="^[+-]?([0-9]+\.?[0-9]*|\.[0-9]+)$"
                title="Favor de introducir un número"
              />
            </MDBCol>
            <MDBCol >
              <MDBInput
                label="test 2"
                v-model="form.oxiSolubre2"
                invalidFeedback="Llena este campo"
                validFeedback="¡Listo!"
                validationEvent="input"
                required
                pattern="^[+-]?([0-9]+\.?[0-9]*|\.[0-9]+)$"
                title="Favor de introducir un número"
              />
            </MDBCol>
            <MDBCol >
              <MDBInput
                label="test 3"
                v-model="form.oxiSolubre3"
                invalidFeedback="Llena este campo"
                validFeedback="¡Listo!"
                validationEvent="input"
                required
                pattern="^[+-]?([0-9]+\.?[0-9]*|\.[0-9]+)$"
                title="Favor de introducir un número"
              >
              </MDBInput>
            </MDBCol>
          </MDBRow>
          <MDBRow> 
              <MDBCol >
                <p> <b>Parámetros químicos</b> </p>
              </MDBCol>
          </MDBRow>
          <MDBRow> 
            <MDBCol >
              <p> pH </p>
            </MDBCol>
            <MDBCol >
              <MDBInput
                label="test 1"
                v-model="form.pH1"
                invalidFeedback="Llena este campo"
                validFeedback="¡Listo!"
                validationEvent="input"
                required
                pattern="^[+-]?([0-9]+\.?[0-9]*|\.[0-9]+)$"
                title="Favor de introducir un número"
              />
            </MDBCol>
            <MDBCol >
              <MDBInput
                label="test 2"
                v-model="form.pH2"
                invalidFeedback="Llena este campo"
                validFeedback="¡Listo!"
                validationEvent="input"
                required
                pattern="^[+-]?([0-9]+\.?[0-9]*|\.[0-9]+)$"
                title="Favor de introducir un número"
              />
            </MDBCol>
            <MDBCol >
              <MDBInput
                label="test 3"
                v-model="form.pH3"
                invalidFeedback="Llena este campo"
                validFeedback="¡Listo!"
                validationEvent="input"
                required
                pattern="^[+-]?([0-9]+\.?[0-9]*|\.[0-9]+)$"
                title="Favor de introducir un número"
              >
              </MDBInput>
            </MDBCol>
          </MDBRow>
          <MDBRow> 
            <MDBCol >
              <p> Fósforo <i> (mg/l)</i> </p>
            </MDBCol>
            <MDBCol >
              <MDBInput
                label="test 1"
                v-model="form.fosforo1"
                invalidFeedback="Llena este campo"
                validFeedback="¡Listo!"
                validationEvent="input"
                required
                pattern="^[+-]?([0-9]+\.?[0-9]*|\.[0-9]+)$"
                title="Favor de introducir un número"
              />
            </MDBCol>
            <MDBCol >
              <MDBInput
                label="test 2"
                v-model="form.fosforo2"
                invalidFeedback="Llena este campo"
                validFeedback="¡Listo!"
                validationEvent="input"
                required
                pattern="^[+-]?([0-9]+\.?[0-9]*|\.[0-9]+)$"
                title="Favor de introducir un número"
              />
            </MDBCol>
            <MDBCol >
              <MDBInput
                label="test 3"
                v-model="form.fosforo3"
                invalidFeedback="Llena este campo"
                validFeedback="¡Listo!"
                validationEvent="input"
                required
                pattern="^[+-]?([0-9]+\.?[0-9]*|\.[0-9]+)$"
                title="Favor de introducir un número"
              >
              </MDBInput>
            </MDBCol>
          </MDBRow>
          <MDBRow> 
            <MDBCol >
              <p> Nitrito <i> (mg/l)</i> </p>
            </MDBCol>
            <MDBCol >
              <MDBInput
                label="test 1"
                v-model="form.nitrito1"
                invalidFeedback="Llena este campo"
                validFeedback="¡Listo!"
                validationEvent="input"
                required
                pattern="^[+-]?([0-9]+\.?[0-9]*|\.[0-9]+)$"
                title="Favor de introducir un número"
              />
            </MDBCol>
            <MDBCol >
              <MDBInput
                label="test 2"
                v-model="form.nitrito2"
                invalidFeedback="Llena este campo"
                validFeedback="¡Listo!"
                validationEvent="input"
                required
                pattern="^[+-]?([0-9]+\.?[0-9]*|\.[0-9]+)$"
                title="Favor de introducir un número"
              />
            </MDBCol>
            <MDBCol >
              <MDBInput
                label="test 3"
                v-model="form.nitrito3"
                invalidFeedback="Llena este campo"
                validFeedback="¡Listo!"
                validationEvent="input"
                required
                pattern="^[+-]?([0-9]+\.?[0-9]*|\.[0-9]+)$"
                title="Favor de introducir un número"
              >
              </MDBInput>
            </MDBCol>
          </MDBRow>
          <MDBRow> 
            <MDBCol >
              <p> Nitrato <i> (mg/l)</i> </p>
            </MDBCol>
            <MDBCol >
              <MDBInput
                label="test 1"
                v-model="form.nitrato1"
                invalidFeedback="Llena este campo"
                validFeedback="¡Listo!"
                validationEvent="input"
                required
                pattern="^[+-]?([0-9]+\.?[0-9]*|\.[0-9]+)$"
                title="Favor de introducir un número"
              />
            </MDBCol>
            <MDBCol >
              <MDBInput
                label="test 2"
                v-model="form.nitrato2"
                invalidFeedback="Llena este campo"
                validFeedback="¡Listo!"
                validationEvent="input"
                required
                pattern="^[+-]?([0-9]+\.?[0-9]*|\.[0-9]+)$"
                title="Favor de introducir un número"
              />
            </MDBCol>
            <MDBCol >
              <MDBInput
                label="test 3"
                v-model="form.nitrato3"
                invalidFeedback="Llena este campo"
                validFeedback="¡Listo!"
                validationEvent="input"
                required
                pattern="^[+-]?([0-9]+\.?[0-9]*|\.[0-9]+)$"
                title="Favor de introducir un número"
              >
              </MDBInput>
            </MDBCol>
          </MDBRow>
          <MDBRow> 
            <MDBCol >
              <p> Amonio <i> (mg/l)</i> </p>
            </MDBCol>
            <MDBCol >
              <MDBInput
                label="test 1"
                v-model="form.amonio1"
                invalidFeedback="Llena este campo"
                validFeedback="¡Listo!"
                validationEvent="input"
                required
                pattern="^[+-]?([0-9]+\.?[0-9]*|\.[0-9]+)$"
                title="Favor de introducir un número"
              />
            </MDBCol>
            <MDBCol >
              <MDBInput
                label="test 2"
                v-model="form.amonio2"
                invalidFeedback="Llena este campo"
                validFeedback="¡Listo!"
                validationEvent="input"
                required
                pattern="^[+-]?([0-9]+\.?[0-9]*|\.[0-9]+)$"
                title="Favor de introducir un número"
              />
            </MDBCol>
            <MDBCol >
              <MDBInput
                label="test 3"
                v-model="form.amonio3"
                invalidFeedback="Llena este campo"
                validFeedback="¡Listo!"
                validationEvent="input"
                required
                pattern="^[+-]?([0-9]+\.?[0-9]*|\.[0-9]+)$"
                title="Favor de introducir un número"
              >
              </MDBInput>
            </MDBCol>
          </MDBRow>
           <MDBRow> 
              <MDBCol >
                <p> <b>Sustrato inorgánico (% de composición en el área muestreada)</b> </p>
              </MDBCol>
          </MDBRow>
          <MDBRow> 
            <MDBCol >
              <p> Rocas <i>(>256 mm) </i> </p>
            </MDBCol>
            <MDBCol >
              <MDBInput
                label="test 1"
                v-model="form.rocas1"
                invalidFeedback="Llena este campo"
                validFeedback="¡Listo!"
                validationEvent="input"
                required
                pattern="^[+-]?([0-9]+\.?[0-9]*|\.[0-9]+)$"
                title="Favor de introducir un número"
              />
            </MDBCol>
            <MDBCol >
              <MDBInput
                label="test 2"
                v-model="form.rocas2"
                invalidFeedback="Llena este campo"
                validFeedback="¡Listo!"
                validationEvent="input"
                required
                pattern="^[+-]?([0-9]+\.?[0-9]*|\.[0-9]+)$"
                title="Favor de introducir un número"
              />
            </MDBCol>
            <MDBCol >
              <MDBInput
                label="test 3"
                v-model="form.rocas3"
                invalidFeedback="Llena este campo"
                validFeedback="¡Listo!"
                validationEvent="input"
                required
                pattern="^[+-]?([0-9]+\.?[0-9]*|\.[0-9]+)$"
                title="Favor de introducir un número"
              >
              </MDBInput>
            </MDBCol>
          </MDBRow>
          <MDBRow> 
            <MDBCol >
              <p> Canto <i>  (64-256 mm)</i> </p>
            </MDBCol>
            <MDBCol >
              <MDBInput
                label="test 1"
                v-model="form.canto1"
                invalidFeedback="Llena este campo"
                validFeedback="¡Listo!"
                validationEvent="input"
                required
                pattern="^[+-]?([0-9]+\.?[0-9]*|\.[0-9]+)$"
                title="Favor de introducir un número"
              />
            </MDBCol>
            <MDBCol >
              <MDBInput
                label="test 2"
                v-model="form.canto2"
                invalidFeedback="Llena este campo"
                validFeedback="¡Listo!"
                validationEvent="input"
                required
                pattern="^[+-]?([0-9]+\.?[0-9]*|\.[0-9]+)$"
                title="Favor de introducir un número"
              />
            </MDBCol>
            <MDBCol >
              <MDBInput
                label="test 3"
                v-model="form.canto3"
                invalidFeedback="Llena este campo"
                validFeedback="¡Listo!"
                validationEvent="input"
                required
                pattern="^[+-]?([0-9]+\.?[0-9]*|\.[0-9]+)$"
                title="Favor de introducir un número"
              >
              </MDBInput>
            </MDBCol>
          </MDBRow>
          <MDBRow> 
            <MDBCol >
              <p> Grava  <i> (2-64 mm) </i> </p>
            </MDBCol>
            <MDBCol >
              <MDBInput
                label="test 1"
                v-model="form.grava1"
                invalidFeedback="Llena este campo"
                validFeedback="¡Listo!"
                validationEvent="input"
                required
                pattern="^[+-]?([0-9]+\.?[0-9]*|\.[0-9]+)$"
                title="Favor de introducir un número"
              />
            </MDBCol>
            <MDBCol >
              <MDBInput
                label="test 2"
                v-model="form.grava2"
                invalidFeedback="Llena este campo"
                validFeedback="¡Listo!"
                validationEvent="input"
                required
                pattern="^[+-]?([0-9]+\.?[0-9]*|\.[0-9]+)$"
                title="Favor de introducir un número"
              />
            </MDBCol>
            <MDBCol >
              <MDBInput
                label="test 3"
                v-model="form.grava3"
                invalidFeedback="Llena este campo"
                validFeedback="¡Listo!"
                validationEvent="input"
                required
                pattern="^[+-]?([0-9]+\.?[0-9]*|\.[0-9]+)$"
                title="Favor de introducir un número"
              >
              </MDBInput>
            </MDBCol>
          </MDBRow>
          <MDBRow> 
            <MDBCol >
              <p> Arena  <i> (0.06-2 mm)</i> </p>
            </MDBCol>
            <MDBCol >
              <MDBInput
                label="test 1"
                v-model="form.arena1"
                invalidFeedback="Llena este campo"
                validFeedback="¡Listo!"
                validationEvent="input"
                required
                pattern="^[+-]?([0-9]+\.?[0-9]*|\.[0-9]+)$"
                title="Favor de introducir un número"
              />
            </MDBCol>
            <MDBCol >
              <MDBInput
                label="test 2"
                v-model="form.arena2"
                invalidFeedback="Llena este campo"
                validFeedback="¡Listo!"
                validationEvent="input"
                required
                pattern="^[+-]?([0-9]+\.?[0-9]*|\.[0-9]+)$"
                title="Favor de introducir un número"
              />
            </MDBCol>
            <MDBCol >
              <MDBInput
                label="test 3"
                v-model="form.arena3"
                invalidFeedback="Llena este campo"
                validFeedback="¡Listo!"
                validationEvent="input"
                required
                pattern="^[+-]?([0-9]+\.?[0-9]*|\.[0-9]+)$"
                title="Favor de introducir un número"
              >
              </MDBInput>
            </MDBCol>
          </MDBRow>
          <MDBRow> 
            <MDBCol >
              <p> Arcilla  <i> (0.004 mm)</i> </p>
            </MDBCol>
            <MDBCol >
              <MDBInput
                label="test 1"
                v-model="form.arcilla1"
                invalidFeedback="Llena este campo"
                validFeedback="¡Listo!"
                validationEvent="input"
                required
                pattern="^[+-]?([0-9]+\.?[0-9]*|\.[0-9]+)$"
                title="Favor de introducir un número"
              />
            </MDBCol>
            <MDBCol >
              <MDBInput
                label="test 2"
                v-model="form.arcilla2"
                invalidFeedback="Llena este campo"
                validFeedback="¡Listo!"
                validationEvent="input"
                required
                pattern="^[+-]?([0-9]+\.?[0-9]*|\.[0-9]+)$"
                title="Favor de introducir un número"
              />
            </MDBCol>
            <MDBCol >
              <MDBInput
                label="test 3"
                v-model="form.arcilla3"
                invalidFeedback="Llena este campo"
                validFeedback="¡Listo!"
                validationEvent="input"
                required
                pattern="^[+-]?([0-9]+\.?[0-9]*|\.[0-9]+)$"
                title="Favor de introducir un número"
              >
              </MDBInput>
            </MDBCol>
          </MDBRow>
          <MDBRow> <br> <MDBBtn color="primary" type="submit" >Generar</MDBBtn> </MDBRow>
          
        </MDBRow>

        <br>
        <!-- Imagen de ejemplo de llenado -->
        <h1 class="display-6">Cómo llenar la ficha del protocolo </h1>
        </MDBTabPane>
        
      </MDBTabContent>
      <!-- Tabs content -->
    </MDBTabs>
  </div>
</div>

</template>

<style>
  * {
    box-sizing: border-box;
  }

  .row {
    padding-top: 50px;
    display: flex;
  }

  /* Create two columns that sits next to each other */
  .column {
    float: left;
    width: 25%;
    padding: 10px;
    padding-top: 25px;
  }
  .column-2 {
    float: left;
    width: 75%;
    padding: 40px;
    }

  /* Clear floats after the columns */
  .row:after {
    content: "";
    display: table;
    clear: both;
  }
  /* Responsive layout - makes the two columns stack on top of each other instead of next to each other */
  @media screen and (max-width: 600px) {
    .column {
      width: 100%;
    }
    .column-2 {
      width: 100%;
    }
  }

  #inner {
    display: table;
    margin: 0 auto;
  }


</style>

<script>
  import { MDBListGroup, 
           MDBListGroupItem,
           MDBTabs,
           MDBTabNav,
           MDBTabContent,
           MDBTabItem,
           MDBTabPane,
           MDBBtn,MDBRow,
           MDBCol,
           MDBInput,
           MDBBadge,
           MDBTable} from "mdb-vue-ui-kit";
  import { ref } from 'vue';
  import router from '../router'
  import axios from 'axios';
  export default {
    components: {
      MDBListGroup,
      MDBListGroupItem,
      MDBTabs,
      MDBTabNav,
      MDBTabContent,
      MDBTabItem,
      MDBTabPane,
      MDBBtn,
      MDBRow,
      MDBCol,
      MDBInput,
      MDBBadge,
      MDBTable
    },
    setup() {
      const checkForm = (e) => {
        e.target.classList.add('was-validated');
      };
      const activeTabId = ref('info-deqi');
      return {
        activeTabId,
        checkForm,

        form: {
          temperatura1: null,
          temperatura2: null,
          temperatura3: null,
          conductividad1: null,
          conductividad2: null,
          conductividad3: null,
          oxiDisuelto1: null,
          oxiDisuelto2: null,
          oxiDisuelto3: null,
          oxiSolubre1: null,
          oxiSolubre2: null,
          oxiSolubre3: null,
          pH1: null,
          pH2: null,
          pH3: null,
          fosforo1: null,
          fosforo2: null,
          fosforo3: null,
          nitrito1: null,
          nitrito2: null,
          nitrito3: null,
          nitrato1: null,
          nitrato2: null,
          nitrato3: null,
          amonio1: null,
          amonio2: null,
          amonio3: null,
          rocas1: null,
          rocas2: null,
          rocas3: null,
          canto1: null,
          canto2: null,
          canto3: null,
          grava1: null,
          grava2: null,
          grava3: null,
          arena1: null,
          arena2: null,
          arena3: null,
          arcilla1: null,
          arcilla2: null,
          arcilla3: null
        } 
      };
    },
    methods: {
      navega: function(url) {
        router.push(url).catch(()=>{});
        },
      // stingify(){
      //   console.log(JSON.stringify(this.form))
      // },
      genera(event){
        event.preventDefault()
        console.log("se llega aqui")

        axios.post(`/api/physchem/export/excel`, JSON.stringify(this.form)  ,{ 
          headers: {
            'Content-Type': 'application/json',
            Accept: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
          }, responseType : 'blob'}
          ).then(resp => {
            const url = window.URL.createObjectURL(new Blob([resp.data],
              { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" })); //excel file
            const link = document.createElement('a');
            link.href = url;
            link.setAttribute('download', 'ProtocoloFisicoQuimica.xlsx'); 
            document.body.appendChild(link);
            link.click();

          });
        console.log("tambien aqui :/") 
        
      }
    }
    
  };
</script>
