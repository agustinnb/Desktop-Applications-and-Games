using System.Collections;
using System.Collections.Generic;
using UnityEngine;
using Prime31;




public class Player_Movement : MonoBehaviour
{

    public float SpeedMove;
    public Rigidbody2D rb;
    public Animator animator;
    private Vector3 movement;
    private CharacterController2D controller;
    
    private void Start()
    {
        animator = GetComponent<Animator>();
        rb = GetComponent<Rigidbody2D>();
    }
    private void Awake()
    {
        controller = GetComponent<CharacterController2D>();
    }

    // Update is called once per frame
    void Update()
    {

        movement = Vector3.zero;
        movement.x = Input.GetAxisRaw("Horizontal");
        movement.y = Input.GetAxisRaw("Vertical");
        if (movement != Vector3.zero)
        {
            MoveCharacter();
            animator.SetFloat("Horizontal", movement.x);
            animator.SetFloat("Vertical", movement.y);
            animator.SetBool("IsMoving", true);
        }
        else
        {
            animator.SetBool("IsMoving", false);



        }
    }

    private void MoveCharacter()
    {
       var translate= (movement * SpeedMove * Time.fixedDeltaTime);
        controller.move(translate);

      
    }

    private void FixedUpdate()
    {
      
    }
    
}
