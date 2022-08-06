<template>

<div class="row">
  <div class="column" style="background-color: #e3f2fd">
    <MDBListGroup  id="inner" >
      <MDBListGroupItem tag="button" @click="navega('/protocolos-info')" action >Protocolos</MDBListGroupItem>
      <MDBListGroupItem tag="button" @click="navega('/protocolo-fisicoquimica')"> Calidad fisicoquímica </MDBListGroupItem>
      <MDBListGroupItem tag="button" action active> HYQI </MDBListGroupItem>
      <MDBListGroupItem tag="button" @click="navega('/cuenca-vegetacion')" action>MEQI</MDBListGroupItem>
      <MDBListGroupItem tag="button" @click="navega('/cuenca-vegetacion')" action>DEQI</MDBListGroupItem>
      <MDBListGroupItem tag="button" @click="navega('/cuenca-conservacion')" action>BEMEQI </MDBListGroupItem>
      <MDBListGroupItem tag="button" @click="navega('/cuenca-jurisdiccion')" action >Material de trabajo</MDBListGroupItem>
    </MDBListGroup>

  </div>

  <div class="column-2" style="background-color:#F8F8F8;">
    <MDBTabs v-model="activeTabId">
      <!-- Tabs navs -->
      <MDBTabNav pills justify tabsClasses="mb-3">
        <MDBTabItem tabId="info-hyqi" href="info-hyqi">Información acerca del protocolo</MDBTabItem>
        <MDBTabItem tabId="calculo-hyqi" href="calculo-hyqi">Cálculo  de calidad fisicoquimica </MDBTabItem>
      </MDBTabNav>
      <!-- Tabs navs -->
      <!-- Tabs content -->
      <MDBTabContent>
        <!-- info protocolo-->
        <MDBTabPane tabId="info-hyqi">
          <p align="justify" > El artículo 3, fracción XLVII, de la Ley de Aguas Nacionales, define como “Ribera o Zona Federal” a las fajas de diez metros de anchura contiguas al cauce de las corrientes o al vaso de los depósitos de propiedad nacional, medidas horizontalmente a partir del nivel de aguas máximas ordinarias. </p>
          <p align="justify">El ecosistema de ribera es un elemento importante en el funcionamiento de un río, sus dimensiones laterales y afluentes, así como muchas de sus dimensiones verticales, están contenidas en su hábitat. Provee servicios a las poblaciones como el suministro de agua, para consumo humano y para el desarrollo de agricultura, ganadería, industria; purificación y transporte de agua, sedimentos y nutrientes; control de inundaciones y sequías; hábitat para las especies; retención de suelo y valores escénicos, culturales y de recreación </p>  
          <h1 class="display-6">Calidad Hidromorfológica</h1>
          <p align = "justify"> El protocolo HYQI (por sus siglas en inglés HYdromorphological Quality Index) incluye datos generales sobre el sitio de muestreo; registro de parámetros físicos y químicos del agua; porcentaje de la composición del sustrato; tipo de bosque del sitio de muestreo, forma de vida predominante. El protocolo contiene doce parámetros hidromorfológicos a evaluar, divididos en tres secciones (las tablas provistas ayudan al reconocimiento de cada uno de los parámetros descritos). </p>
          <p align="justify"  class="fs-4">Sección 1: La calidad de la cuenca</p>
          <p align="justify" >  Se examina la cobertura vegetal, la estabilidad del banco, talla del sustrato y desarrollo de ganadería y agricultura en la zona de ribera:</p>
          <MDBTable borderless responsive>
            <thead class="table-info">
              <tr>
                <th scope="col">Banco estable </th>
                <th scope="col">Banco parcialmente erosionado</th>
                <th scope="col">Banco totalmente erosionado</th>
              </tr>
            </thead>
            <tbody>
              <tr>
                <td > <img src="../assets/banco_estable_hyqi.jpg" alt="banco estable" 
                  style="width: 163px; height: 126px;" class="rounded" /> </td>
                <td><img src="../assets/banco_parc_estable.jpg" alt="banco parcialmente estable" 
                  style="width: 160px; height: 126px;" class="rounded" /></td>
                <td><img src="../assets/banco_tot_erosionado.jpg" alt="banco erosionado" 
                  style="width: 160px; height: 126px;" class="rounded" /></td>
              </tr>
            </tbody>
          </MDBTable>
          <p align="justify" > Las características del sustrato relaciona la heterogeneidad del sustrato orgánico e inorgánico y define el número de microambientes potenciales en los que se pueden desarrollar organismos acuáticos:</p>
          <MDBTable borderless responsive>
            <thead class="table-info">
              <tr>
                <th scope="col">Sustratos inorgánicos (arena, grava, roca emergente) </th>
                <th scope="col">Sustratos orgánicos (algas, vegetación acuática)</th>
              </tr>
            </thead>
            <tbody>
              <tr>
                <td > <img src="../assets/sust_inorg.jpg" alt="inorganico" 
                  style="width: 196px; height: 129px;" class="rounded" /> </td>
                <td><img src="../assets/sust_organ.jpg" alt="organico" 
                  style="width: 196px; height: 129px;" class="rounded" /></td>
              </tr>
            </tbody>
          </MDBTable>
          <p align="justify"> Las principales intervenciones humanas en los ríos son por la presencia de trucheros, actividades agrícolas y ganaderas:</p>
          <MDBTable borderless responsive>
            <thead class="table-info">
              <tr>
                <th scope="col">Sin presencia de cultivos o zonas para el ganado </th>
                <th scope="col">+80% de suelo para uso piscícola, agrícola y ganadero</th>
              </tr>
            </thead>
            <tbody>
              <tr>
                <td > <img src="../assets/sin_ganaderia.jpg" alt="buen suelo" 
                  style="width: 196px; height: 129px;" class="rounded" /> </td>
                <td><img src="../assets/con_ganaderia.jpg" alt="suelo usado" 
                  style="width: 196px; height: 129px;" class="rounded" /></td>
              </tr>
            </tbody>
          </MDBTable>

          <p align="justify" class="fs-4">Sección 2: La hidrología del río</p>
          <p align="justify" >  Se examina la hidrología según la naturalidad del canal, se tienen en cuenta las interrupciones del flujo por la construcción de presas.</p>
          <MDBTable borderless responsive>
            <thead class="table-info">
              <tr>
                <th scope="col">Sin presencia de  presas</th>
                <th scope="col">Presencia de presas</th>
              </tr>
            </thead>
            <tbody>
              <tr>
                <td > <img src="../assets/sin_presas.jpg" alt="sin presa" 
                  style="width: 196px; height: 129px;" class="rounded" /> </td>
                <td><img src="../assets/con_presas.jpg" alt="con presa" 
                  style="width: 196px; height: 129px;" class="rounded" /></td>
              </tr>
            </tbody>
          </MDBTable>
          <p align="justify" >El régimen de velocidad y profundidad se determina observando la presencia de cuatro combinaciones posibles en el lecho del río: rápido-somero, rápido-profundo, lento-somero y lento-profundo. Una zona somera es aquella con profundidades menores a 0.4 <i>m</i> y una zona rápida está por encima de los 0.3 <i>m/s</i> </p>
          <img src="../assets/reg_velprof.jpg" height="220" width="570" class="img-fluid  hover-shadow" alt="ejemplo aforo" />  <br><br>
          <p align="justify"> La alteración del canal se puede reconocer por la construcción de infraestructura aledaña al margen de ambos lados del río.</p>
          <MDBTable borderless responsive>
            <thead class="table-info">
              <tr>
                <th scope="col">Sin alteraciones en el canal</th>
                <th scope="col">Alteraciones parciales</th>
                <th scope="col">Canal totalmente modificado</th>
              </tr>
            </thead>
            <tbody>
              <tr>
                <td > <img src="../assets/sin_alteracion.jpg" alt="nada alterado" 
                  style="width: 163px; height: 126px;" class="rounded" /> </td>
                <td><img src="../assets/alteracion_parc.jpg" alt="Alteraciones parciales" 
                  style="width: 160px; height: 126px;" class="rounded" /></td>
                  <td><img src="../assets/alteracion_total.jpg" alt="todo alterado" 
                  style="width: 160px; height: 126px;" class="rounded" /></td>
              </tr>
            </tbody>
          </MDBTable>
          <p align="justify"> El estado del canal se reconoce por la cantidad del agua que circula y el nivel que ocupa en ambos bancos de la orilla. </p>
          <MDBTable borderless responsive>
            <thead class="table-info">
              <tr>
                <th scope="col">El agua llega a la base de los bancos, sustrato no expuesto</th>
                <th scope="col">Muy poca agua en el canal</th>
              </tr>
            </thead>
            <tbody>
              <tr>
                <td > <img src="../assets/buen_estado.jpg" alt="sin presa" 
                  style="width: 196px; height: 129px;" class="rounded" /> </td>
                <td><img src="../assets/mal_estado.jpg" alt="con presa" 
                  style="width: 196px; height: 129px;" class="rounded" /></td>
              </tr>
            </tbody>
          </MDBTable>

          <p align="justify" class="fs-4">Sección 3: Las perturbaciones antropogénicas</p>
          <p align="justify" > Estas tienen que ver con la contaminación del agua por efluentes urbanos, desarrollo humano y contaminación.</p>
          <MDBTable borderless responsive>
            <thead class="table-info">
              <tr>
                <th scope="col">Sin efluentes al río</th>
                <th scope="col">Descarga directa de residuos al río</th>
              </tr>
            </thead>
            <tbody>
              <tr>
                <td > <img src="../assets/sin_efluentes.jpg" alt="sin efluentes" 
                  style="width: 196px; height: 129px;" class="rounded" /> </td>
                <td><img src="../assets/descarga_eflu.jpg" alt="descargas al rio" 
                  style="width: 196px; height: 129px;" class="rounded" /></td>
              </tr>
            </tbody>
          </MDBTable>
          <p align="justify">El desarrollo de infraestructura urbana refiere a la presencia de caminos, carreteras o puentes asociadas al canal del río, incluyendo extracción en este.</p>
          <MDBTable borderless responsive>
            <thead class="table-info">
              <tr>
                <th scope="col">Sin infraestructura</th>
                <th scope="col">Carreteras y construcciones</th>
              </tr>
            </thead>
            <tbody>
              <tr>
                <td > <img src="../assets/sin_construcc.jpg" alt="sin construiir" 
                  style="width: 196px; height: 129px;" class="rounded" /> </td>
                <td><img src="../assets/construcciones.jpg" alt="construcciones" 
                  style="width: 196px; height: 129px;" class="rounded" /></td>
              </tr>
            </tbody>
          </MDBTable>
          <MDBTable borderless responsive>
            <thead class="table-info">
              <tr>
                <th scope="col">Sin extracción</th>
                <th scope="col">Extracción directa en la ribera</th>
              </tr>
            </thead>
            <tbody>
              <tr>
                <td > <img src="../assets/no_extracc.jpg" alt="sin extracción" 
                  style="width: 196px; height: 129px;" class="rounded" /> </td>
                <td><img src="../assets/extracc_directa.jpg" alt="extracción directa" 
                  style="width: 196px; height: 129px;" class="rounded" /></td>
              </tr>
            </tbody>
          </MDBTable>
          <p align="justify">Se examina también la contaminación de residuos sólidos por basura y escombros que son vertidos al lecho del río.</p>
          <MDBTable borderless responsive>
            <thead class="table-info">
              <tr>
                <th scope="col">Sin basura, y sin escombros</th>
                <th scope="col">Presencia difusa de basura y escombros</th>
                <th scope="col">Presencia abundante de basura y escombros</th>
              </tr>
            </thead>
            <tbody>
              <tr>
                <td > <img src="../assets/sin_basura.jpg" alt="nada basura" 
                  style="width: 163px; height: 135px;" class="rounded" /> </td>
                <td><img src="../assets/poca_basura.jpg" alt="poca basura" 
                  style="width: 160px; height: 135px;" class="rounded" /></td>
                  <td><img src="../assets/mucha_basura.jpg" alt="basura y escombros" 
                  style="width: 160px; height: 135px;" class="rounded" /></td>
              </tr>
            </tbody>
          </MDBTable>
          <p align="justify" class="fs-3"> &emsp; &emsp; Evaluación</p>
          <p align="justify" >El protocolo HYQI tiene un total de 120 puntos y cada sección tiene un total de 40 puntos. El puntaje parcial permitirá identificar la sección mejor conservada de las 3 descritas anteriormente. 
          El puntaje final es el resultado de la suma de los doce parámetros analizados y se verifica su estado de acuerdo al puntaje final y una tabla de colores. Si el valor supera los 85 puntos está en condición de ser considerado como sitio de referencia por su buena calidad hidromorfológica. </p>
          <MDBTable border="dark">
            <thead >
              <tr>
                <th scope="col">Calidad hidromorfológica</th>
                <th scope="col">Puntaje</th>
              </tr>
            </thead>
            <tbody>
              <tr class="table-primary">
                <td > Óptima </td>
                <td> 120-85</td>
              </tr>
              <tr class="table-success">
                <td > Media </td>
                <td> 84-47</td>
              </tr>
              <tr class="table-warning">
                <td > Mala </td>
                <td> 46-13</td>
              </tr>
              <tr class="table-danger">
                <td > Pobre </td>
                <td> &lt;12 </td>
              </tr>
            </tbody>
          </MDBTable>

        </MDBTabPane>

        <!-- Calculo con microservicio-->
        <MDBTabPane tabId="calculo-hyqi">
        <h1 class="display-5">Cálculo del protocolo</h1>
        <br> <p align="justify"> Aquí puedes introducir los parámetros hidromorfológicos del agua obtenidos en la colecta, hasta abajo encontrarás un ejemplo del llenado de esta ficha. Al dar clic en Generar se descargará un archivo editable de Excel donde puedes visualizar los datos, y se incluyen celdas listas para calcular también el aforo.</p>
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
      const activeTabId = ref('info-hyqi');
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
