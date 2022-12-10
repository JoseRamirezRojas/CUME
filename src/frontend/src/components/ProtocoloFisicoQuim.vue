<template>

<div class="row">
  <div class="column" style="background-color: #e3f2fd">
    <MDBListGroup  id="inner" >
      <MDBListGroupItem tag="button" @click="navega('/protocolos-info')" action >Protocolos</MDBListGroupItem>
      <MDBListGroupItem tag="button" action active> Calidad fisicoquímica </MDBListGroupItem>
      <MDBListGroupItem tag="button" @click="navega('/protocolo-hyqi')" action > HYQI </MDBListGroupItem>
      <MDBListGroupItem tag="button" @click="navega('/protocolo-deqi')" action>DEQI</MDBListGroupItem>
      <MDBListGroupItem tag="button" @click="navega('/protocolo-meqi')" action>MEQI</MDBListGroupItem>
      <MDBListGroupItem tag="button" @click="navega('/protocolo-bemeqi')" action>BEMEQI </MDBListGroupItem>
      <MDBListGroupItem tag="button" @click="navega('/cuenca-jurisdiccion')" action >Material de trabajo</MDBListGroupItem>
    </MDBListGroup>

  </div>

  <div class="column-2" style="background-color:#F8F8F8;">
    <MDBTabs v-model="activeTabId">
      <!-- Tabs navs -->
      <MDBTabNav pills justify tabsClasses="mb-3">
        <MDBTabItem tabId="info-fisicoquim" href="info-fisicoquim">Información acerca del protocolo</MDBTabItem>
        <MDBTabItem tabId="calculo-fisicoquim" href="calculo-fisicoquim">Cálculo  de calidad fisicoquimica </MDBTabItem>
      </MDBTabNav>
      <!-- Tabs navs -->
      <!-- Tabs content -->
      <MDBTabContent>
        <!-- info protocolo-->
        <MDBTabPane tabId="info-fisicoquim">
          <p align="justify" > De acuerdo con las características fisicoquímicas, los ríos de la cuenca de México se definen como arroyos de montaña de una región subtropical, agua templada entre 5 - 17 °C, moderadamente ácida con pH 5.5 - 8, y poco mineralizada (34.5-175 μS/cm). La mayoría de las sitios evaluados presentan un caudal permanente con variable aforo (0.01-1.5 ㎥/s), que se puede atribuir a la estacionalidad, aunque en la parte media y baja de las cuencas, se debe también a la extracción in situ de estructuras de regulación hídrica. El nitrógeno inorgánico disuelto y el fósforo reactivo soluble coinciden con lugares donde se permiten actividades humanas con usos directos del agua del río. </p>
          <br>
          <h1 class="display-6">Calidad Fisicoquímica</h1>
          <img src="../assets/fisicoquim_1.jpg" width="237" class="img-fluid  hover-shadow" alt="..." style="float:right;
          margin-left: 20px" />
          <p align="justify">La temperatura del agua es el parámetro que determina gran parte de los procesos fisicoquímicos y biológicos en el agua. Siempre es importante medirla e intentar determinar su relación con las distintas épocas del año, la altitud del río y las posibles acciones humanas que pueden modificarla (por ejemplo represamientos donde se almacena y calienta el agua). La temperatura del agua siempre se tiene que registrar ya que puede modificarse en unos cuantos centímetros a partir del nacimiento de agua o en zonas donde el río acumula agua y registra más de un metro de profundidad. Esta se registra con termómetros o sensores acoplados a dataloggers. </p>  
          <br>
          <p align = "justify"> La velocidad de corriente (V) y el aforo (Q), expresado en ㎥/s, se registran con un medidor de corriente y representa el tiempo que tarda en recorrer un objeto una distancia conocida, es decir:</p>
          <vue-mathjax :formula="formulaTime" /> donde <i> d </i> = distancia y <i> t </i> = tiempo.
          <p align="justify">  La estimación del aforo del río permite comparar la cantidad de
            agua que pasa en el sitio de colecta en distinta época del año (m<sup>2</sup>). Se recomienda registrar los valores en el mismo punto en cada momento y se usa la siguiente fórmula: </p>
          <vue-mathjax :formula="formulaCapacity" /> Con <i> A</i> = área y <i>V</i> = velocidad.
          <br> <br> <p align="justify"> Para las orillas calculamos un triángulo, y entonces </p>
          <vue-mathjax :formula="formulaTriangle" /> donde <i>B</i> = distancia entre la orilla y primera medida del ancho del cauce,
            y <i>h</i> = la primera profundidad a partir de la orilla. 
           <br> <br> <p align="justify"> Para el resto de las áreas se calculan considerando un polígono cuadrangular y entonces </p>
           <vue-mathjax :formula="formulaTrapeze" /> donde <i>B</i> = 1<sup>a</sup> profundidad,  <i>b</i> = 2<sup>a</sup> profundidad,
            y <i>h</i> = distancia del intervalo. 
          <p align="justify"> Cada área calculada la multiplicamos por su velocidad de corriente para obtener aforos parciales y obtenemos el aforo total calculando la suma de los aforos parciales.</p>
          <vue-mathjax :formula="formulaTotalCapacity" /> <br>
          <img src="../assets/fig_aforo.png" height="400" width="610" class="img-fluid  hover-shadow" alt="ejemplo aforo" />
          <br> <br>
          <p align="justify"> En esta figura tenemos el perfil de un río en el que se calcula el aforo. El área de los extremos es un triángulo y las áreas intemedias son polígonos cuadrangulares.</p>
          <br>
          <h1 class="display-6">Nutrientes</h1>
          <p align="justify"> En general, la cuenca de México presenta una baja concentración de micronutrientes en el nacimiento de los ríos y se puede ir incrementando a lo largo del cauce. </p>
          <p align="justify"> La concentración de micronutrientes en el agua es el parámetro que puede definir el su uso para el humano. Generalmente la concentración de fósforo y nitrógeno total evidencían la limitación (sistemas oligotróficos) o el excedente por aportes externos al río (sistemas eutróficos). </p>
          <p align="justify"> Como referencia, en la siguiente tabla se encuentran los límites sugeridos según varias obras para la clasificación trófica de ríos y relación de uso del agua: </p>
          <MDBTable borderless responsive>
            <thead class="table-info">
              <tr>
                <th scope="col">Parámetro (unidades)</th>
                <th scope="col">Límite oligotrófico-mesotrófico</th>
                <th scope="col">Límite mesotrófico-eutrófico</th>
              </tr>
            </thead>
            <tbody>
              <tr  class="table-active">
                <th  scope="row">Doods (1998) &emsp; &emsp;  </th>    <td></td> <td></td>
                
              </tr>
              <tr>
                <td scope="row">Nitrógeno total <i>(μg/l)</i> </td>
                <td>700</td>
                <td>1500</td>
              </tr>
              <tr>
                <td scope="row">Fósforo total <i>(μg/l)</i></td>
                <td>25</td>
                <td>75</td>
              </tr>
              <tr class="table-active">
                <th scope="row">Vida acuática, DOF (2012) &emsp; &emsp; </th>    <td></td> <td></td>
              </tr>
              <tr>
                <td scope="row">Nitrógeno total <i>(mg/l) </i><sub>(promedio mensual) </sub> </td>
                <td>15</td>
              </tr>
              <tr>
                <td scope="row">Fósforo total <i>(mg/l)</i> <sub>(promedio mensual) </sub></td>
                <td>5</td>
              </tr>
              <tr  class="table-active">
                <th scope="row">Uso público e infiltración, DOF (1997) &emsp; &emsp; </th>    <td></td> <td></td>
              </tr>
              <tr>
                <td scope="row">Nitrógeno total <i> (mg/l)</i></td>
                <td>40</td>
              </tr>
              <tr>
                <td scope="row">Fósforo total <i>(mg/l)</i> </td>
                <td>20</td>
              </tr>
              <tr  class="table-active">
                <th scope="row">Uso y consumo humano, DOF (2012) &emsp; &emsp; </th>    <td></td> <td></td>
              </tr>
              <tr>
                <td scope="row">Nitratos <i> (mg/l)</i>  </td>
                <td>10</td>
              </tr>
              <tr>
                <td scope="row">Nitritos <i> (mg/l) </i></td>
                <td>0.05</td>
              </tr>
              <tr>
                <td scope="row">Nitrógeno amoniacal <i>(mg/l) </i></td>
                <td>0.5</td>
              </tr>
            </tbody>
          </MDBTable>
          <p align="justify">El registro de las formas particuladas e inorgánicas del nitrógeno (nitrito, nitrato y amonio) e inorgánicas del fósforo (ortofosfatos), determinan la predominancia de los grupos algales que lo pueden incorporar por metabolismos particulares. O bien, determinar las fuentes de contaminación de compuestos químicos inorgánicos por aguas residuales urbanas o agroquímicos. </p>
          <p align="justify"> También se pueden evaluar diversos macronutrientes para determinar la disponibilidad de algunos elementos inorgánicos que requieren algunos organismos para su desarrollo metabólico o estructura celular. Por ejemplo, carbonatos y sulfatos que se requieren para realizar la fotosíntesis y la fabricación de mucílago extracelular. </p>

          
        </MDBTabPane>

        <!-- Calculo con microservicio-->
        <MDBTabPane tabId="calculo-fisicoquim">
        <h1 class="display-5">Cálculo del protocolo</h1>
        <br> <p align="justify"> Aquí puedes introducir los parámetros fisicoquímicos del agua obtenidos en la colecta. Al dar clic en Generar se descargará un archivo editable de Excel donde puedes visualizar los datos, y se incluyen celdas listas para calcular también el aforo.</p>
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
           MDBTable,
           MDBBtn,MDBRow,
           MDBCol,
           MDBInput } from "mdb-vue-ui-kit";
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
      MDBTable,
      MDBBtn,
      MDBRow,
      MDBCol,
      MDBInput
    },
    setup() {
      const checkForm = (e) => {
        e.target.classList.add('was-validated');
      };
      const activeTabId = ref('info-fisicoquim');
      return {
        activeTabId,
        checkForm,
        formulaTime: "$$V = \\frac {d}{t}$$",
        formulaCapacity: "$$Q = A \\cdot V$$",
        formulaTriangle: "$$A = \\frac {B \\cdot h}{2}$$",
        formulaTrapeze: "$$A= \\frac{(B+b)h}{2}$$",
        formulaTotalCapacity: "$$\\sum Q_i = \\sum (A_i \\cdot V_i)$$",
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
