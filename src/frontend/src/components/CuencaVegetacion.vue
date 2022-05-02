<template>

<div class="row">
  <div class="column" style="background-color: #e3f2fd">
    <MDBListGroup  id="inner" >
      <MDBListGroupItem tag="button" @click="navega('/cuenca-mexico')" action>La cuenca</MDBListGroupItem>
      <MDBListGroupItem tag="button" @click="navega('/cuenca-hidrologia')" action >Hidrología </MDBListGroupItem>
      <MDBListGroupItem tag="button" @click="navega('/cuenca-geologia')" action >Geología </MDBListGroupItem>
      <MDBListGroupItem tag="button" action active>Clima y vegetación</MDBListGroupItem>
      <MDBListGroupItem tag="button" @click="navega('/cuenca-conservacion')" action>Áreas de conservación </MDBListGroupItem>
      <MDBListGroupItem tag="button" @click="navega('/cuenca-jurisdiccion')" action >Gestión y marco jurídico</MDBListGroupItem>
    </MDBListGroup>

  </div>
  <div class="column-2 " style="background-color:#F8F8F8;" >
    <h2>Clima y vegetación</h2> <br>
    <p class="fs-5" align="justify" >La vegetación de árboles y matorrales en la Cuenca de México está representada principalmente por las siguientes comunidades: </p>

    <MDBRow :cols="['1','md-2']" class="g-3">
      <MDBCol>
        <MDBCard >
          <MDBCardImg src="https://www.lifeder.com/wp-content/uploads/2020/05/bosques-de-mexico.jpg" top alt="pinus" style="max-height: 270px"/>
          <MDBCardHeader> <strong>Bosque de pino (Pinus spp.)</strong> </MDBCardHeader>
              <MDBCardBody>
                <MDBCardText class="text-start">
                  Típico de los ecosistemas de montaña, húmedos o subhúmedos; y crecen entre los 2,350 y 4,000 m de altitud, con lluvias anuales de entre 700 y 1,200 mm.
                </MDBCardText>
              </MDBCardBody>
        </MDBCard>
      </MDBCol>
      <MDBCol>
        <MDBCard >
          <MDBCardImg src="https://upload.wikimedia.org/wikipedia/commons/8/8e/Pino-encino.jpg" top alt="quercus" style="max-height: 270px"/>
          <MDBCardHeader> <strong>Bosque de encino (Quercus spp.)</strong></MDBCardHeader>
              <MDBCardBody>
                <MDBCardText class="text-start">
                  Se encuentran entre los 300 y 3,000 m, con lluvias anuales. Con frecuencia bosques de pino y encino crecen juntos formando comunidades mixtas. Los encinos tienen entre 5 y 12 m de altura. 
                </MDBCardText>
              </MDBCardBody>
        </MDBCard>
      </MDBCol>
      <MDBCol>
        <MDBCard>
          <MDBCardImg src="https://www.researchgate.net/profile/Sergio-Zamudio/publication/259800886/figure/fig4/AS:341168874377224@1458352406993/Figura-10-Enebro-Juniperus-flaccida-cerca-de-Palomas-municipio-de-Atarjea-Guanajuato.png" top alt="enebro" style="max-height: 270px"/>
          <MDBCardHeader><strong>Bosque de enebros (Juniperus spp.)</strong></MDBCardHeader>
          <MDBCardBody>          
            <MDBCardText class="text-start">
              Son arbustos o árboles de menos de 6m de altura, forman bosques bajos con abundante vegetación herbácea. Son comunes en las partes norte, este y noreste de la cuenca, entre los 2,400 y los 2,800 m de altitud. Se desarrollan sobre laderas o planicies semiáridas, con lluvias anuales de entre 600 y 800 mm. 
            </MDBCardText>
          </MDBCardBody>
        </MDBCard>
      </MDBCol>
      <MDBCol>
        <MDBCard>
          <MDBCardImg src="https://encrypted-tbn0.gstatic.com/images?q=tbn:ANd9GcQoGt1sfbTPS5Vvqpa3mmXJSQK6J4SrHVktNA&usqp=CAU" top alt="quercus mycro" style="max-height: 270px"/>
          <MDBCardHeader> <strong>Matorrales de encinos chaparros (Quercus microphylla)</strong> </MDBCardHeader>
          <MDBCardBody>          
            <MDBCardText class="text-start">
              Esta comunidad forma una cubierta densa de arbustos bajos (40 a 80 cm de altura). Este tipo de bosque se encuentra principalmente al noreste de la cuenca, en áreas semiáridas con 700 y 900 mm de lluvia media anual.  
            </MDBCardText>
          </MDBCardBody>
        </MDBCard>
      </MDBCol>

      <MDBCol>
        <MDBCard>
          <MDBCardImg src="data:image/jpeg;base64,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" top alt="enebro" style="max-height: 270px"/>
          <MDBCardHeader> <strong>Matorrales xerófilos.</strong> </MDBCardHeader>
          <MDBCardBody>          
            <MDBCardText class="text-start">
              Este tipo de vegetación está compuesta por varias comunidades arbustivas y son frecuentes en la parte norte de la cuenca, donde las precipitaciones son más pobres. En general ocupan partes bajas de la cuenca, entre 2,250 y 2,700 m de altitud, en áreas de precipitación media anual inferior a los 700 mm. En la Sierra de Guadalupe quedan restos del matorral  Eisenhardtia polystachya, que está desapareciendo por la presión antropogénica. 
            </MDBCardText>
          </MDBCardBody>
        </MDBCard>
      </MDBCol>
      <MDBCol>
        <MDBCard>
          <MDBCardImg src="https://mapio.net/images-p/43043335.jpg" top alt="halofila" style="max-height: 270px"/>
          <MDBCardHeader> <strong>Vegetación halófila</strong> </MDBCardHeader>
          <MDBCardBody>          
            <MDBCardText class="text-start">
              Esta vegetación domina algunas partes más bajas de la cuenca, sobre los lechos de los antiguos lagos. Es particularmente común el lecho seco del antiguo lago de Texcoco. Este tipo de vegetación ocupaba los márgenes del algo, pero con la construcción del drenaje de la cuenca y el secado de los lagos, se ha extendido a las partes más bajas. 
            </MDBCardText>
          </MDBCardBody>
        </MDBCard>
      </MDBCol>

    </MDBRow>
  <br> <br>
  <p class="fs-5" align="justify" >Y las principales comunidades de pastizales presentes en la cuenca son las siguientes: </p>
    <MDBRow :cols="['1','md-2']" class="g-3">
      <MDBCol>
        <MDBCard>
          <MDBCardImg src="https://encrypted-tbn0.gstatic.com/images?q=tbn:ANd9GcSAuYAM6vo1GOzbfuhBuhwIC0X6GFnHtxQo8g&usqp=CAU" top alt="espiga" style="max-height: 270px"/>
          <MDBCardHeader> <strong>Espiga negra (Hilaria cenchroides)</strong> </MDBCardHeader>
          <MDBCardBody>          
            <MDBCardText class="text-start">
              Es la principal formación de pastizal de la cuenca de México, son comunes al pie de la Sierra Nevada. Entre los 2,300 y 2,700 m de altitud, y con precipitaciones anuales de 600 y 700 mm.
            </MDBCardText>
          </MDBCardBody>
        </MDBCard>
      </MDBCol>
      <MDBCol>
        <MDBCard>
          <MDBCardImg src="data:image/jpeg;base64,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" top alt="espiga" style="max-height: 270px"/>
          <MDBCardHeader> <strong>Gramíneas anuales  </strong> </MDBCardHeader>
          <MDBCardBody>          
            <MDBCardText class="text-start">
              En las planicies del centro y norte de la cuenca, a 2,300 y 2,400 m de altitud y sobre áreas fuertemente perturbadas, se desarrolla una comunidad en la que dominan dos pastizales, Aristida adscensionis y Bouteloua simplex.
            </MDBCardText>
          </MDBCardBody>
        </MDBCard>
      </MDBCol>

      <MDBCol>
        <MDBCard>
          <MDBCardImg src="https://encrypted-tbn0.gstatic.com/images?q=tbn:ANd9GcRI3G3wRH2eRKa_8lcLwSSJUv1-Ss7sII70iA&usqp=CAU "  top alt="espiga" style="max-height: 270px"/>
          <MDBCardHeader> <strong>Zacatones (Muhlenbergia y Festuca)  </strong> </MDBCardHeader>
          <MDBCardBody>          
            <MDBCardText class="text-start">
              A más de los 4,000 m de altitud por encima del bosque de Pinus hartwegii, se encuentran los pastizales de Muhlenbergia y Festuca. Estos pastos o zacatones, forman matas erectas de 60 a 120 cm de altura.  
            </MDBCardText>
          </MDBCardBody>
        </MDBCard>
      </MDBCol>
    </MDBRow>
    
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
  padding: 10px;
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
  import {MDBListGroup, MDBListGroupItem , MDBRow, MDBCol, MDBCard, MDBCardImg, MDBCardBody, MDBCardHeader, MDBCardText } from "mdb-vue-ui-kit";
  import router from '../router'
  export default {
    components: {
      MDBListGroup,
      MDBListGroupItem,
      MDBRow,
      MDBCol,
      MDBCard,
      MDBCardBody,
      MDBCardHeader,
      MDBCardText,
      MDBCardImg
    }, 
    methods: {
    navega: function(url) {
      router.push(url).catch(()=>{});
      },
    }
  };
</script>