using System.Collections;
using System.Collections.Generic;
using UnityEngine;


public class Camera_Follow : MonoBehaviour
{
    private Transform playerTransform;
    public float smoothing;

    // Start is called before the first frame update
    void Start()
    {
        playerTransform = GameObject.FindGameObjectWithTag("Player").transform;



    }
    private void LateUpdate()
    {
        Vector3 playerPosition = new Vector3(playerTransform.position.x, playerTransform.position.y, transform.position.z);
        if (transform.position != playerTransform.position)
        {
            transform.position = Vector3.Lerp(transform.position, playerPosition, smoothing);
            
        }
  
        //      Vector3 temp = transform.position;

        //       temp.x = playerTransform.position.x;
        //  temp.y = playerTransform.position.y;
        //     transform.position = temp;
    }

    // Update is called once per frame
    void Update()
    {
        
    }

}
