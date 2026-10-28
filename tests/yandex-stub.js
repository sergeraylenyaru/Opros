/* Local adapter test double; no real Yandex key or network traffic. */
(function(){
 class Events{constructor(){this.handlers={};}add(name,fn){(this.handlers[name]??=[]).push(fn);}emit(name,values){(this.handlers[name]||[]).forEach(fn=>fn({get:key=>values[key]}));}}
 class Placemark{constructor(coords){this.events=new Events();this.geometry={getCoordinates:()=>this.coords,setCoordinates:value=>{this.coords=value;}};this.coords=coords;window.testYandexPlacemark=this;}}
 class Map{constructor(el,state){this.center=state.center;this.zoom=state.zoom;this.events=new Events();this.geoObjects={add(){}};this.search={events:new Events(),getResult:()=>Promise.resolve({geometry:{getCoordinates:()=>[59.93,30.31]}})};this.controls={get:()=>this.search};window.testYandexMap=this;}getCenter(){return this.center;}getZoom(){return this.zoom;}setCenter(coords,zoom){this.center=coords;this.zoom=zoom;}destroy(){this.destroyed=true;}}
 window.ymaps={Map,Placemark,ready:fn=>fn()};
})();
