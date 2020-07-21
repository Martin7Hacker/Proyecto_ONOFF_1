/*
 * 
 *  Ensender led en Arduino
 *  Driver para Arduino
 *  Encender 5 Led´s
 *  Autor Martin Grasso.
 *  
 */

 const int MaxLED = 5;
 const int MinLED = 0;
 int led[MaxLED] = {2,3,4,5,6}; 
 int dato;

 void setup(){
      for(int i=MinLED;i<=MaxLED;i++){  
        pinMode(led[i],OUTPUT);
        Serial.begin(9600);
      }
 }
 
void loop(){
     dato=Serial.read();
 switch (dato) {
   /****************/
    case '0':
       digitalWrite(led[0],HIGH);
      break;
     
    case '1':
       digitalWrite(led[0],LOW);
      break;
    /****************/

    /****************/
    case '2':
       digitalWrite(led[1],HIGH);
      break;
   
    case '3':
       digitalWrite(led[1],LOW);
      break;
    /****************/

/****************/
    case '4':
       digitalWrite(led[2],HIGH);
      break;
     
    case '5':
       digitalWrite(led[2],LOW);
      break;
    /****************/

    /****************/
    case '6':
       digitalWrite(led[3],HIGH);
      break;
   
    case '7':
       digitalWrite(led[3],LOW);
      break;
    /****************/

    /****************/
    case '8':
       digitalWrite(led[4],HIGH);
      break;
   
    case '9':
       digitalWrite(led[4],LOW);
      break;
    /****************/
    }
  delay(50);
}



