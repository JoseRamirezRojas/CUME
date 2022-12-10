<template>

<div class="row">
  <div class="column" style="background-color: #e3f2fd">
    <MDBListGroup  id="inner" >
      <MDBListGroupItem tag="button" @click="navega('/protocolos-info')" action >Protocolos</MDBListGroupItem>
      <MDBListGroupItem tag="button" @click="navega('/protocolo-fisicoquimica')" action> Calidad fisicoquímica </MDBListGroupItem>
      <MDBListGroupItem tag="button" @click="navega('/protocolo-hyqi')" action > HYQI </MDBListGroupItem>
      <MDBListGroupItem tag="button" @click="navega('/protocolo-deqi')" action>DEQI</MDBListGroupItem>
      <MDBListGroupItem tag="button" action active>MEQI</MDBListGroupItem>
      <MDBListGroupItem tag="button" @click="navega('/protocolo-bemeqi')" action>BEMEQI </MDBListGroupItem>
      <MDBListGroupItem tag="button" @click="navega('/cuenca-jurisdiccion')" action >Material de trabajo</MDBListGroupItem>
    </MDBListGroup>

  </div>

  <div class="column-2" style="background-color:#F8F8F8;">
    <MDBTabs v-model="activeTabId">
      <!-- Tabs navs -->
      <MDBTabNav pills justify tabsClasses="mb-3">
        <MDBTabItem tabId="info-meqi" href="info-meqi"      >Información del protocolo</MDBTabItem>
        <MDBTabItem tabId="calculo-meqi" href="calculo-meqi">Cálculo de índice MEQI</MDBTabItem>
      </MDBTabNav>
      <!-- Tabs navs -->
      <!-- Tabs content -->
      <MDBTabContent>
        <!-- info protocolo-->
        <MDBTabPane tabId="info-meqi">
          <h1 class="display-6">Macroalgas </h1>
          <p align="justify" > El término macroalga bentónica se aplica a las algas visibles a simple vista que viven asociadas a cualquier sustrato sólido en los ambientes lóticos, entre las que se encuentran <i>Cyanobacterias</i>, <i>Rhodophyta</i>, <i>Chlorophyta</i> y <i>Heterokontophyta</i>.  <br> En México, las macroalgas se limitan en muchos casos al ámbito de la investigación y éstos todavía no se han incluido en las redes de control de calidad gestionadas por la <a href="https://www.miteco.gob.es/es/agua/temas/planificacion-hidrologica/marco-del-agua/default.aspx" target="_blank">Administración.</a>  </p>
          <img src="../assets/macroalga_ilust.png"  class="img-fluid  hover-shadow" alt="meqi" style="float:right; height: 207px; width: 253px; margin-left: 20px" />
          <p align="justify" > De acuerdo con la DMA, las macroalgas se consideran útiles para la detección y seguimiento de las presiones fisicoquímicas e hidromorfológicas de los ríos, entre las cuales se incluyen:</p>
          <ul style = "text-align: justify">
            <li align="justify">Reducción de la transparencia del agua.</li>
            <li align="justify">Eutrofía</li>
            <li align="justify">Variaciones del régimen del caudal, en la continuidad del río y características morfológicas del lecho. </li>
          </ul>
          <br>
          <p align = "justify"> Las macroalgas son sensibles a los cambios de calidad fisicoquímica y tienen un tiempo de respuesta de semanas o meses. La desaparición de una especie en el sistema puede reflejar un daño altamente significativo relacionado con el volúmen y la calidad de agua. Es importante resaltar que su abundancia puede estar influida por variaciones anuales o interanuales, por lo que es necesario analizarla durante un periodo de tiempo de varias temporadas. Para empezar, lo recomendable es tener representadas las estaciones climatológicas más contrastantes del año, por ejemplo, primavera, verano, otoño e invierno. O bien, época de lluvias, secas frías y secas cálidas.  </p>
          <h1 class="display-6">Colecta </h1>
          <p align = "justify"> Para su registro, se recomienda establecer un transecto de 10 metros de longitud en el tramo de río a evaluar y donde sea posible establecer diferentes zonas que abarcan tipos de sustratos y regímenes de velocidad de corriente y profundidad. En el transecto se registra la abundancia relativa ó porcentaje de cobertura algal calculada en una unidad de muestreo (círculo de 20 cm de diámetro). Este método de colecta se puede apreciar en la siguiente fotografía:</p>
          <img src="../assets/transecto_meqi.png"  class="img-fluid  hover-shadow" alt="transecto" style=" width: 349px; " /> <br> <br>
          <p align="justify" >  El número de unidades puede variar (3-10) de acuerdo a los crecimientos que se encuentren. La cobertura de cada crecimiento se acompaña del registro de variables microambientales (velocidad de corriente, profundidad, talla del sustrato e iluminación). </p>
          <h1 class="display-6">Análisis</h1>
          <img src="../assets/analisis_meqi.jpg"  class="img-fluid  hover-shadow" alt="diatomea" style="float:left; height: 133px; margin-right: 15px" />
          <p align="justify" > De cada unidad de muestreo se recolectan tres a cinco talos o un crecimiento abundante, se colocan en frascos y se transportan en una hielera al laboratorio para su posterior identificación taxonómica, para la cual puede consultarse principalmente <a href="https://link.springer.com/article/10.1007/s10750-009-0074-5" target="_blank" rel="noopener noreferrer"> Bojorge-García <i>et al.</i> (2010) </a>, <a href="https://www.scielo.org.mx/scielo.php?pid=S2007-33802019000100128&script=sci_arttext_plus&tlng=en" target="_blank" rel="noopener noreferrer"> Carmona-Jiménez <i>et al.</i> (2019) </a>,  entre otros recursos más.</p>
          <h1 class="display-6">Evaluación</h1>
          <p align="justify" > La cobertura porcentual de cada especie se transcribe en el formato que se ubica en la pestaña de CÁLCULO DE ÍNDICE MEQI. Y esta cobertura porcentual (<i>h</i>) para cada especie por transecto es igual al promedio de su cobertura en los 10 cuadrantes definidos para recolectar.</p> 
          <p align="justify"> Por la naturaleza de la colecta para este protocolo, si el sitio se colecta en más de una ocasión, se numera en orden progresivo y así, el valor de la calidad ecológica se calcula obteniendo el promedio de todas las fechas de colecta. </p> 
          <p align="justify"> Los valores indicadores de las 20 especies de macroalgas reportadas para la cuenca fueron determinadas por <a href="http://revista.ib.unam.mx/index.php/bio/article/view/3899" target="_blank" rel="noopener noreferrer"> Carmona-Jiménez <i>et al.</i> ( 2021 ) </a> y se pueden consultar en esta tabla, o en la hoja de cálculo que puede descargarse en la pestaña CÁLCULO DE ÍNDICE MEQI </p> <br>
          <MDBTable borderless responsive>
            <thead class="table-info">
              <tr>
                <th scope="col">Taxa</th>
                <th scope="col">Valor indicador (<i>v </i>)</th>
              </tr>
            </thead>
            <tbody>
              <tr  class="table-active">
                <th  scope="row">Cyanobacteria &emsp; &emsp; &emsp; &emsp;  </th>    <td></td> 
                
              </tr>
              <tr>
                <td scope="row"><i>Calothrix</i> sp. </td>
                <td>1</td>
              </tr>
              <tr>
                <td scope="row"><i>Coleodesmium wrangelii</i> Borzí ex Geitler</td>
                <td>1</td>
              </tr>
              <tr>
                <td scope="row"><i>Nostoc aff. parmelioides</i> Kützing ex Bornet &amp; Flahault</td>
                <td>1</td>
              </tr>
              <tr>
                <td scope="row"><i>Phormidium aff. autumnale</i> Gomont</td>
                <td>5</td>
              </tr>
              <tr>
                <td scope="row"><i>Placoma regularis</i> Brody &amp; Ingerfeld</td>
                <td>1</td>
              </tr>
              <tr class="table-active">
                <th scope="row">Rhodophyta &emsp; &emsp; &emsp; &emsp; </th>   <td></td>
              </tr>
              <tr>
                <td scope="row"><i>Batrachospermum gelatinosum</i> (Linnaeus) De Candolle</td>
                <td>1</td>
              </tr>
              <tr>
                <td scope="row"><i>Paralamanea mexicana</i> (Kützing) Vis &amp; Sheath</td>
                <td>1</td>
              </tr>
              <tr>
                <td scope="row"><i> 'Chantrasia' </i> stage of undefined Rhodophyta</td>
                <td>5</td>
              </tr>
              <tr class="table-active">
                <th scope="row">Chlorophyta &emsp; &emsp; &emsp; &emsp; </th>  <td></td>
              </tr>
              <tr>
                <td scope="row"><i>Ulothrix</i> sp. </td>
                <td>1</td>
              </tr>
              <tr>
                <td scope="row"><i>Prasiola mexicana</i> Agardh </td>
                <td>1</td>
              </tr>
              <tr>
                <td scope="row"><i>Spirogyra</i> sp. </td>
                <td>3</td>
              </tr>
              <tr>
                <td scope="row"><i>Mougeotia</i> sp. </td>
                <td>3</td>
              </tr>
              <tr>
                <td scope="row"><i>Cladophora</i> sp. </td>
                <td>3</td>
              </tr>
              <tr>
                <td scope="row"><i>Rhizoclonium</i> sp. </td>
                <td>3</td>
              </tr>
              <tr>
                <td scope="row"><i>Oedogonium</i> sp. </td>
                <td>3</td>
              </tr>
              <tr>
                <td scope="row"><i>Tetraspora gelatinosa</i> (Vaucher) Desvaux </td>
                <td>3</td>
              </tr>
              <tr>
                <td scope="row"><i>Draparnaldia mutabilis</i> (Roth) Bory </td>
                <td>3</td>
              </tr>
              <tr>
                <td scope="row"><i>Stigeoclonium tenue</i> (Agardh) Kützing </td>
                <td>5</td>
              </tr>
              <tr  class="table-active">
                <th scope="row">Xantophyta &emsp; &emsp; &emsp; &emsp; </th>    <td></td> 
              </tr>
              <tr>
                <td scope="row"><i>Tribonema</i> sp. </td>
                <td>5</td>
              </tr>
              <tr>
                <td scope="row"><i>Vaucheria bursata</i> (Müller) Agardh </td>
                <td>1</td>
              </tr>
            </tbody>
          </MDBTable>
          <br> <br> 
          <p align="justify"> Los valores indicadores (<i>v</i>) y la abundancia de las macroalgas (<i>h</i>) reportados en la muestra nos llevan a calcular el índice MEQI que ,de manera análoga a DEQI, se calcula a partir de la suma de los productos del valor indicador (<i>v)</i> de cada especie por su abundancia según el conteo de cada taxón (<i>h</i>), dividida entre la suma de la abundancia de todos los taxones. Lo anterior se resume en la siguiente fórmula: </p>
          <vue-mathjax :formula="formula" />  <br>
          <p align="justify">El valor del MEQI se clasifica en cinco categorías definidas como: alta, buena, moderada, pobre y mala, donde los valores inferiores a 1.5 pueden considerarse sitios de referencia de buena calidad ecológica y los valores superiores a 3.51 indican sitios perturbados. Las clases intermedias indican especies que pueden tolerar mayores concentraciones de nutrientes o alteraciones en la calidad hidromorfológica.</p>
          <p align="justify"> Las especies que se repiten en distintos ensambles pueden tolerar perturbaciones en la calidad hidromorfológica y/o fisicoquímica del agua. En estos casos, es importante determinar su preferencia en cada clase con base en su abundancia. </p>
          <MDBTable border="gray" bordered>
            <thead >
              <tr>
                <th scope="col">Calidad ecológica</th>
                <th scope="col">MEQI</th>
              </tr>
            </thead>
            <tbody>
              <tr class="table-primary">
                <td > Alta </td>
                <td> 1 - 1.5</td>
              </tr>
              <tr class="table-success">
                <td > Buena </td>
                <td> 1.51 - 2.5</td>
              </tr>
              <tr class="table-warning">
                <td > Moderada </td>
                <td> 2.51 - 3.5</td>
              </tr>
              <tr class="table-danger">
                <td > Pobre </td>
                <td> 3.51 - 4.5 </td>
              </tr>
              <tr class="bg-danger p-2 bg-opacity-25">
                <td > Mala </td>
                <td> 4.51 - 5 </td>
              </tr>
            </tbody>
          </MDBTable>
        </MDBTabPane>

        <!-- Calculo con microservicio-->
        <MDBTabPane tabId="calculo-meqi">
        <h1 class="display-5">Cálculo del protocolo</h1>
        <br> <p align="justify"> Aquí puedes descargar el formato para el cálculo del índice y llenarlo con datos del conteo de macroinvertebrados obtenidos en la colecta, debajo del botón encontrarás información de cómo llenar adecuadamente el formato. </p>
        <div class="d-grid gap-2">
          <MDBBtn color="primary" download @click="genera">
            Generar 
          </MDBBtn> 
        </div>
        <!-- Imagen de ejemplo de llenado -->
          <br><h1 class="display-6">Hoja 1: MEQI</h1>
          <img src="../assets/hoja1MEQI.png"  class="img-fluid  hover-shadow" alt="hoja1" style=" height: 252px; " /> <br> <br>
          <p align="justify">En esta hoja, hasta arriba, deberás anotar los datos del proyecto para el que se realiza este cálculo, los nombres de los autores y revisor, y las fechas de colecta y revisión. </p>
          <p align="justify">Posteriormente según los datos analizados de la colecta deberás anotar el conteo de cada taxón reconocido, según su acomodo por filas.</p>
          <p align="justify">Cada columna, a partir de la columna D, representa un sitio de colecta involucrado en el proyecto. De inicio se incluyen 20 sitios numerados pero se puede cambiar el nombre de las columnas por el nombre del sitio recolectada, o bien, asignar claves a cada sitio y ponerlas en dichas columnas. </p>
          <p align="justify">Al momento de escribir los datos de conteo de especies la hoja hará los cálculos actualizados al momento, desplegando el valor MEQI de cada sitio en la parte de abajo de la hoja en las filas 30-32 con los códigos de colores descritos en la pestaña INFORMACIÓN DEL PROTOCOLO en este sitio; y hasta arriba de la hoja se mostrará también el valor MEQI promedio de todos los sitios que tengan datos introducidos, así como informar la calidad ecológica en la que se pueden clasificar los sitos introducidos.  </p>
          <br><h1 class="display-6">Hoja 2: Ficha de colecta</h1>
          <img src="../assets/hoja2MEQI.png"  class="img-fluid  hover-shadow" alt="hoja2" style=" height: 252px; " /> <br> <br>
          <p align="justify">En esta hoja, se encuentra disponible la ficha para llevar control de la colecta, esta en teoría deberá llevarse en formato impreso durante la colecta, para luego capturar la información correspondiente de manera digital.</p>
          <p align="justify">Se tendrá que cambiar el valor de "Cuadrante X" por el número de cuadrante que se esté colectando en el sitio, y tras anotar los valores de velocidad de corriente e irradiancia el promedio de estos se calculará en la ficha digital.</p>
          
        </MDBTabPane>
      </MDBTabContent>
      <!-- Tabs content -->
    </MDBTabs>
  </div>
</div>

</template>

<style scoped>
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

<script >
  import { MDBListGroup, 
           MDBListGroupItem,
           MDBTabs,
           MDBTabNav,
           MDBTabContent,
           MDBTabItem,
           MDBTabPane,
           MDBTable,
           MDBBtn,
          } from "mdb-vue-ui-kit";
  import { ref } from 'vue';
  import router from '../router'
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
    },
    setup() {
      const activeTabId = ref('info-meqi');
     
      return {
        activeTabId,
        formula: '$$MEQI = \\frac{ \\sum_{i=1}^n(v_i \\cdot h_i )} { \\sum_{i=1}^n h_i} $$',
      };
    },
    methods: {
      navega: function(url) {
        router.push(url).catch(()=>{});
        },
      genera(){
        console.log("se llega aqui")
        const link = document.createElement('a');
        link.href = './assets/files/ProtocoloMEQI.xlsx';
        link.setAttribute('download', 'ProtocoloMEQI.xlsx'); 
        document.body.appendChild(link);
        link.click();
      },
      
    } 
  };
</script>
